using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Threading;
using System.Xml;
using static classStandardRoutines;
using static ProcToEmail.classProcToEmail;

namespace ProcToEmail
{
    class classWork
    {
        private QueueParameters udsQueueParameters;
        private SortedDictionary<string, int> sdRetryList;
        public struct QueueParameters
        {
            public string ProcessName;
            public string QueueName;
            public string LogName;
            public string LogDirectory;
            public string ArchiveFolder;
            public string ErrorFolder;
            public string DatabaseType;
            public string ConnectionString;
            public string StoredProcedure;
            public string HTMLTemplate;
            public int AutoParameters;
            public string EmailField;
            public string CCAddress;
            public string BCCAddress;
            public string IDField;
            public string FromAddress;
            public string FromName;
            public string SubjectField;
            public string DebugEmail;
            public int DebugMode;
            public string SMTPServer;
            //public int WaitTime;
            public List<ParameterMapping> ParameterMappings;
            public List<InputParameter> InputParameters;

        }

        public struct ParameterMapping
        {

            public string DBField;
            public string ReplaceVar;
        }

        public struct InputParameter
        {

            public string Name;
            public string Value;
        }

        public void Begin(string sProcessName, string sQueueData, ref SortedDictionary<string, int> sdRetryList)
        {
            this.sdRetryList = sdRetryList;
            XmlDocument xdOptionsDoc;
            xdOptionsDoc = new XmlDocument();
            udsQueueParameters.ProcessName = sProcessName;
            xdOptionsDoc.LoadXml(sQueueData);
            udsQueueParameters.QueueName = GetSingleNode_Str(xdOptionsDoc, ("//@Name"));
            udsQueueParameters.LogDirectory = FmtDirectoryName(GetSingleNode_Str(xdOptionsDoc, ("//LogDirectory")));
            udsQueueParameters.LogName = sProcessName + "_" + udsQueueParameters.QueueName;
            udsQueueParameters.ArchiveFolder = FmtDirectoryName(GetSingleNode_Str(xdOptionsDoc, ("//ArchiveFolder")));
            udsQueueParameters.ErrorFolder = FmtDirectoryName(GetSingleNode_Str(xdOptionsDoc, ("//ErrorFolder")));
            udsQueueParameters.DatabaseType = GetSingleNode_Str(xdOptionsDoc, ("//DatabaseType"));
            udsQueueParameters.ConnectionString = GetSingleNode_Str(xdOptionsDoc, ("//ConnectionString"));
            udsQueueParameters.StoredProcedure = GetSingleNode_Str(xdOptionsDoc, ("//StoredProcedure"));
            udsQueueParameters.HTMLTemplate = GetSingleNode_Str(xdOptionsDoc, ("//HTMLTemplate"));
            udsQueueParameters.AutoParameters = GetSingleNode_Int(xdOptionsDoc, ("//AutoParameters"));
            udsQueueParameters.EmailField = GetSingleNode_Str(xdOptionsDoc, ("//EmailField"));
            udsQueueParameters.IDField = GetSingleNode_Str(xdOptionsDoc, ("//IDField"));
            udsQueueParameters.CCAddress = GetSingleNode_Str(xdOptionsDoc, ("//CCAddress"));
            udsQueueParameters.BCCAddress = GetSingleNode_Str(xdOptionsDoc, ("//BCCAddress"));
            udsQueueParameters.FromAddress = GetSingleNode_Str(xdOptionsDoc, ("//FromAddress"));
            udsQueueParameters.FromName = GetSingleNode_Str(xdOptionsDoc, ("//FromName"));
            udsQueueParameters.SubjectField = GetSingleNode_Str(xdOptionsDoc, ("//SubjectField"));
            //udsQueueParameters.WaitTime = GetSingleNode_Int(xdOptionsDoc, ("//WaitTime"));
            udsQueueParameters.DebugEmail = GetSingleNode_Str(xdOptionsDoc, ("//DebugEmail"));
            udsQueueParameters.DebugMode = GetSingleNode_Int(xdOptionsDoc, ("//DebugMode"));
            udsQueueParameters.SMTPServer = GetSingleNode_Str(xdOptionsDoc, ("//SMTPServer"));

            if (xdOptionsDoc.SelectSingleNode("//InputParameter/Parameter") != null)
            {
                udsQueueParameters.InputParameters = new List<InputParameter>();
                foreach (XmlNode nodeParam in xdOptionsDoc.SelectNodes("//InputParameter/Parameter"))
                {
                    InputParameter udsParam = new InputParameter();
                    udsParam.Name = GetSingleNode_Str(nodeParam, "@Name");
                    udsParam.Value = nodeParam.InnerText;
                    udsQueueParameters.InputParameters.Add(udsParam);
                }
            }

            if (xdOptionsDoc.SelectSingleNode("//ParameterMapping/DataMap") != null)
            {
                udsQueueParameters.ParameterMappings = new List<ParameterMapping>();     
                foreach (XmlNode nodeParam in xdOptionsDoc.SelectNodes("//ParameterMapping/DataMap"))
                {
                    ParameterMapping udsParam = new ParameterMapping();
                    udsParam.DBField = GetSingleNode_Str(nodeParam, "@DBField");
                    udsParam.ReplaceVar = nodeParam.InnerText;
                    udsQueueParameters.ParameterMappings.Add(udsParam);
                }
            }
            BuildDirectoryTree(udsQueueParameters.LogDirectory);
        }


        public void ProcessQueue(ref bool bStopRequest, ref DateTime dtLastPurge)
        {
            // Prepare for Logging
            CreateLogFile(udsQueueParameters.LogName, udsQueueParameters.LogDirectory);

            LogMessage(udsQueueParameters.LogName, udsQueueParameters.LogDirectory, "Queue: " + udsQueueParameters.QueueName + " started processing");
            try { 
            bool outcome = ProcessRecords(udsQueueParameters);
                if (outcome == true)
                {
                    LogMessage(udsQueueParameters.LogName, udsQueueParameters.LogDirectory, "Queue: " + udsQueueParameters.QueueName + " completed succesfully with no errors");
                } else
                {
                    LogMessage(udsQueueParameters.LogName, udsQueueParameters.LogDirectory, "Queue: " + udsQueueParameters.QueueName + " completed with some errors");
                }
            }
            catch(Exception ex)
            {
                LogMessage(udsQueueParameters.LogName, udsQueueParameters.LogDirectory, "Queue: " + udsQueueParameters.QueueName + " failed with error " + ex.ToString());
            }

            // Wait for 2 minutes so process doesn't run again
            //Thread.Sleep(udsQueueParameters.WaitTime * 1000);
        }
           
    }
}
