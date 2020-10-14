using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Sybase.Data.AseClient;
using System.Data;
using System.Data.SqlClient;
using static classStandardRoutines;

namespace ProcToEmail
{
    class classProcToEmail
    {

        public static bool ProcessRecords(classWork.QueueParameters udsQueueParameters)
        {

            bool outcome = true;
            DataTable dt = new DataTable();

            if (udsQueueParameters.DatabaseType == "Sybase")
            {
                dt = ProcessSybase(udsQueueParameters);
            }

            else if (udsQueueParameters.DatabaseType == "MSSQL")
            {
                dt = ProcessMSSQL(udsQueueParameters);
            }

            foreach (DataRow row in dt.Rows)
            {
                string emailBody = System.IO.File.ReadAllText(udsQueueParameters.HTMLTemplate);
                string emailAddress = row[udsQueueParameters.EmailField].ToString();
                string id = row[udsQueueParameters.IDField].ToString();

                if (udsQueueParameters.AutoParameters == 1)
                {
                    foreach(DataColumn col in dt.Columns)
                    {
                        string replaceVar = "[$" + col.ColumnName.ToString() + "$]";
                        string replacementText = row[col.ColumnName.ToString()].ToString();
                        emailBody = emailBody.Replace(replaceVar, replacementText);
                    }

                }
                else
                {
                    foreach (var replacement in udsQueueParameters.ParameterMappings)
                    {
                        string replacementText = row[replacement.DBField].ToString();
                        emailBody = emailBody.Replace(replacement.ReplaceVar, replacementText);
                    }
                }

                if (!emailAddress.Contains("@"))
                {
                    LogMessage(udsQueueParameters.LogName, udsQueueParameters.LogDirectory, "Record: " + id + " does not have a valid email address");
                    outcome = false;
                }
                else
                {
                    if (udsQueueParameters.DebugMode == 1)
                    {
                        emailAddress = udsQueueParameters.DebugEmail;
                    }

                    SendMail(udsQueueParameters.FromAddress, udsQueueParameters.FromName, emailAddress, udsQueueParameters.CCAddress, udsQueueParameters.BCCAddress, udsQueueParameters.Subject, emailBody, udsQueueParameters.SMTPServer, null);
                    LogMessage(udsQueueParameters.LogName, udsQueueParameters.LogDirectory, "Record: " + id + " successfully sent email to " + emailAddress);
                }
            }

            return outcome;
        }


        private static DataTable ProcessSybase(classWork.QueueParameters udsQueueParameters)
        {

            AseConnection ncsDB = new AseConnection(udsQueueParameters.ConnectionString);
            List<AseParameter> param = new List<AseParameter>();
            foreach (var thisParam in udsQueueParameters.InputParameters)
            {
                string dynamicParam = ParamMapping(thisParam.Value);
                param.Add(new AseParameter(thisParam.Name, dynamicParam));
            }

            DataTable dt = getDataTableFromASEProc(ncsDB, udsQueueParameters.StoredProcedure, param);
            return dt;
        }

        private static DataTable ProcessMSSQL(classWork.QueueParameters udsQueueParameters)
        {

            SqlConnection sqlDB = new SqlConnection(udsQueueParameters.ConnectionString);
            List<SqlParameter> param = new List<SqlParameter>();
            foreach (var thisParam in udsQueueParameters.InputParameters)
            {
                string dynamicParam = ParamMapping(thisParam.Value);
                param.Add(new SqlParameter(thisParam.Name, dynamicParam));
            }

            DataTable dt = getDataTableFromProc(sqlDB, udsQueueParameters.StoredProcedure, param);
            return dt;
        }

        private static DataTable getDataTableFromProc(SqlConnection connString, string sProc, List<SqlParameter> sqlParameters)
        {
            connString.Open();
            DataTable tempDT = new DataTable();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = sProc;
            if (sqlParameters != null)
            {
                foreach (SqlParameter param in sqlParameters)
                    cmd.Parameters.AddWithValue(param.ParameterName, param.Value);
            }
            cmd.Connection = connString;
            var result = cmd.ExecuteReader();
            tempDT.Load(result);
            connString.Close();
            return tempDT;
        }

        private static DataTable getDataTableFromASEProc(AseConnection connString, string sProc, List<AseParameter> sqlParameters)
        {
            connString.Open();
            DataTable tempDT = new DataTable();
            AseCommand cmd = new AseCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = sProc;
            if (sqlParameters != null)
            {
                foreach (AseParameter param in sqlParameters)
                    cmd.Parameters.AddWithValue(param.ParameterName, param.Value);
            }
            cmd.Connection = connString;
            var result = cmd.ExecuteReader();
            tempDT.Load(result);
            connString.Close();
            return tempDT;
        }

        private static void ExecuteNonScaler(SqlConnection connString, string sProc, List<SqlParameter> sqlParameters)
        {

            connString.Open();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = sProc;
            if (sqlParameters != null)
            {
                foreach (SqlParameter param in sqlParameters)
                    cmd.Parameters.AddWithValue(param.ParameterName, param.Value);
            }
            cmd.Connection = connString;
            cmd.ExecuteNonQuery();
            connString.Close();

        }

        private static void ExecuteASENonScaler(AseConnection connString, string sProc, List<AseParameter> sqlParameters)
        {

            connString.Open();
            AseCommand cmd = new AseCommand();
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = sProc;
            if (sqlParameters != null)
            {
                foreach (AseParameter param in sqlParameters)
                    cmd.Parameters.AddWithValue(param.ParameterName, param.Value);
            }
            cmd.Connection = connString;
            cmd.ExecuteNonQuery();
            connString.Close();

        }

        private static string ParamMapping(string param)
        {
            string dynamicParam = param;
            if (dynamicParam == "[$today$]")
            {
                dynamicParam = DateTime.Now.ToString("dd MMM yyyy");
            }
            else if (dynamicParam == "[$yesterday$]")
            {
                dynamicParam = DateTime.Now.AddDays(-1).ToString("dd MMM yyyy");
            }
            else if (dynamicParam == "[$tomorrow$]")
            {
                dynamicParam = DateTime.Now.AddDays(1).ToString("dd MMM yyyy");
            }
            return dynamicParam;
        }
    }
}
