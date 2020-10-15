using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using System.Threading;
using System.Xml;
using static classStandardRoutines;

namespace ProcToEmail
{
    public class classQueue
    {
        private string sServiceName;
        private bool bStopRequest;
        private string sQueueData;
        private string sQueueName;
        private bool bStopped;

        public bool StopRequest
        {
            set { bStopRequest = value; }
        }
        public string ServiceName
        {
            set { sServiceName = value; }
        }
        public string QueueData
        {
            set { sQueueData = value; }
        }
        public bool Stopped
        {
            set { bStopped = value; }
        }
        public string QueueName
        {
            set { sQueueName = value; }
        }

        private DateTime dtLastPurgeCheck = default(DateTime);

        public void Start(object State)
        {
            //string sRunTime;
            //string sRunDay;
            Thread thMain;
            XmlDocument xdOptionsDoc = new XmlDocument();
            string sErrorText = "";
            classWork oWork;
            string[] sParts;
            int iStartHour;
            int iStartMinute;
            List<string> lRunDays = null;
            List<string> lRunTimes = null;
            //List<DateTime> dtEnd = new List<DateTime>();
            SortedDictionary<string, int> sdRetryList = new SortedDictionary<string, int>();

            bStopRequest = false;
            bStopped = false;

            // Extract the Queue Attributes
            try
            {
                xdOptionsDoc.LoadXml(sQueueData);
                thMain = Thread.CurrentThread;
                sQueueName = GetSingleNode_Str(xdOptionsDoc, ("//@Name"));
                if (thMain.Name == null)
                    thMain.Name = sQueueName;
                if (GetSingleNode_Str(xdOptionsDoc, ("//@RunTime")) != "" && GetSingleNode_Str(xdOptionsDoc, ("//@RunTime")) != null)
                {
                    lRunTimes = GetSingleNode_Str(xdOptionsDoc, ("//@RunTime")).Split(',').ToList();
                }
                //if (sRunTime.ToUpper() == "NOW")
                //    sRunTime = DateTime.Now.ToShortTimeString();
                if (GetSingleNode_Str(xdOptionsDoc, ("//@RunDay")) != "" && GetSingleNode_Str(xdOptionsDoc, ("//@RunDay")) != null)
                {
                    lRunDays = GetSingleNode_Str(xdOptionsDoc, ("//@RunDay")).Split(',').ToList();
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry(sServiceName, "Worker thread unable to load Queue attributes from XML", EventLogEntryType.Error);
                bStopped = true;
                return;
            }

            // Log a Start Message
            if (State != "Debug")
                EventLog.WriteEntry(sServiceName, "Queue " + sQueueName + " has started", EventLogEntryType.Information);

            // Keep Processing the Queue until asked to stop
            while (bStopRequest == false)
            {
                bool bTimeinRange = false;
                bool bDayinRange = false;

                // Generate a processing end time [2 minutes after the Run Time]
                if (lRunTimes == null)
                {
                    bTimeinRange = true;
                }
                else
                {
                    foreach (string sRunTime in lRunTimes)
                    {
                        if (ValidTime(sRunTime) == false)
                            throw new ApplicationException("Invalid run time: " + sRunTime);

                        sParts = sRunTime.Split(new[] { ":" }, StringSplitOptions.None);
                        iStartHour = System.Convert.ToInt32(sParts[0]);
                        iStartMinute = System.Convert.ToInt32(sParts[1]);
                        DateTime dtEnd = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, iStartHour, iStartMinute, 0).AddMinutes(2);
                        if (InTimeRange(StringToDate(sRunTime), dtEnd))
                        {
                            bTimeinRange = true;
                        }


                        //dtEnd.Add = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, iStartHour, iStartMinute, 0).AddMinutes(2);
                    }
                }

                if (lRunDays == null)
                {
                    bDayinRange = true;
                }
                else
                {
                    foreach (string sRunDay in lRunDays)
                    {
                        if (DateTime.Now.ToString("dddd") == sRunDay)
                        {
                            bDayinRange = true;
                        }
                    }
                }

                // Determine if it's time to run ...
                if (bDayinRange && bTimeinRange)
                {
                    try
                    {

                        // Process the Queue
                        oWork = new classWork();
                        oWork.Begin(sServiceName, sQueueData, ref sdRetryList);
                        oWork.ProcessQueue(ref bStopRequest, ref dtLastPurgeCheck);
                        oWork = null;
                        Thread.Sleep(120000);
                        //bStopRequest = true;

                    }
                    catch (Exception ex)
                    {
                        sErrorText = "General Processing Error [Queue.Start]: " + "\r\n" + "\r\n";
                        sErrorText += ex.Message;
                        EventLog.WriteEntry(sServiceName, "Queue " + sErrorText, EventLogEntryType.Information);
                    }
                }

                Thread.Sleep(7000);

                if (bStopRequest)
                    bStopped = true;

                if (State == "Debug")
                    break;
            }
        }
    }
}
