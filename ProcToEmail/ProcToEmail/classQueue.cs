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
            string sRunTime;
            Thread thMain;
            XmlDocument xdOptionsDoc = new XmlDocument();
            string sErrorText = "";
            classWork oWork;
            string[] sParts;
            int iStartHour;
            int iStartMinute;
            DateTime dtEnd;
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
                sRunTime = GetSingleNode_Str(xdOptionsDoc, ("//@RunTime"));
                if (sRunTime.ToUpper() == "NOW")
                    sRunTime = DateTime.Now.ToShortTimeString();
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

                // Generate a processing end time [2 minutes after the Run Time]
                if (sRunTime == "")
                    dtEnd = default(DateTime);
                else
                {
                    if (ValidTime(sRunTime) == false)
                        throw new ApplicationException("Invalid run time: " + sRunTime);
                    sParts = sRunTime.Split(new[] { ":" }, StringSplitOptions.None);
                    iStartHour = System.Convert.ToInt32(sParts[0]);
                    iStartMinute = System.Convert.ToInt32(sParts[1]);
                    dtEnd = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, iStartHour, iStartMinute, 0).AddMinutes(2);
                }

                // Determine if it's time to run ...
                if ((sRunTime == "") | (InTimeRange(StringToDate(sRunTime), dtEnd)))
                {
                    try
                    {

                        // Process the Queue
                        oWork = new classWork();
                        oWork.Begin(sServiceName, sQueueData, ref sdRetryList);
                        oWork.ProcessQueue(ref bStopRequest, ref dtLastPurgeCheck);
                        oWork = null;

                        // Pause until the end time has been reached
                        if (sRunTime != "")
                        {
                            //while (DateTime.DateDiff(DateInterval.Second, DateTime.Now, dtEnd) >= 0)
                            while ((DateTime.Now - dtEnd).TotalSeconds >= 0)
                            {
                                Thread.Sleep(3000);
                                if (bStopRequest)
                                    break;
                            }
                        }
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
