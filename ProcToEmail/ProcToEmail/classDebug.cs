using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Reflection;
using System.Xml;
using System.Diagnostics;

namespace ProcToEmail
{
    class ClassDebug
    {
        public static void DebugMain()
        {
            Assembly objAssembly;
            FileInfo fiOptionsFile = null;
            string sProcess;
            classQueue oQueue;
            XmlDocument xdOptionsDoc = new XmlDocument();

            XmlNodeList xnlWorkList;
            DateTime lastruntime = default(DateTime);
            bool stoprequest = false;

            Debug.WriteLine("Start ...");

            // Get the Options File
            try
            {
                objAssembly = Assembly.GetExecutingAssembly();
                fiOptionsFile = new FileInfo(objAssembly.Location);
                xdOptionsDoc.Load(fiOptionsFile.DirectoryName + @"\Options\Options.xml");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Unable to load the Options file" + "\r\n" + ex.Message);
                Debug.WriteLine("===================================================================");
                return;
            }

            // Read each Queue and process
            try
            {
                sProcess = classStandardRoutines.GetSingleNode_Str(xdOptionsDoc, ("/Process/@Name"));
                xnlWorkList = xdOptionsDoc.SelectNodes("//Queues/Queue");
                if (xnlWorkList != null)
                {
                    foreach (XmlNode xnWorkNode in xnlWorkList)
                    {
                        oQueue = new classQueue();
                        oQueue.ServiceName = sProcess.Trim();
                        oQueue.QueueData = xnWorkNode.OuterXml;
                        oQueue.Start("Debug");
                        oQueue = null;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Queue processing error" + "\r\n" + ex.Message);
                Debug.WriteLine("===================================================================");
                return;
            }

            Debug.WriteLine("End ...");
            Debug.WriteLine("===================================================================");
        }

    }
}
