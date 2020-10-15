using System;
using System.IO;
using System.Xml;
using IWshRuntimeLibrary;
using System.Threading;
using System.ComponentModel;
using System.Net.Mail;
using System.Collections;
using System.Collections.Generic;

static class classStandardRoutines
{

    // Fixed and updated

    public const Boolean PURGE_SUBDIRECTORY = true;
    public const Boolean LEAVE_SUBDIRECTORY = false;
    public const Boolean OVERWRITE = true;

    public class clsCompareFileInfo : IComparer
    {
        public int Compare(object x, object y)
        {
            System.IO.FileInfo File1;
            System.IO.FileInfo File2;

            File1 = (System.IO.FileInfo)x;
            File2 = (System.IO.FileInfo)y;

            return DateTime.Compare(File1.LastWriteTime, File2.LastWriteTime);
        }
    }

    public class clsCompareFileInfoName : IComparer
    {
        public int Compare(object x, object y)
        {
            System.IO.FileInfo File1;
            System.IO.FileInfo File2;

            File1 = (System.IO.FileInfo)x;
            File2 = (System.IO.FileInfo)y;

            return string.Compare(File1.Name, File2.Name);
        }
    }

    public class ComboData
    {
        public object Value;
        public string Description;

        public ComboData(object NewValue, string NewDescription)
        {
            Value = NewValue;
            Description = NewDescription;
        }

        public override string ToString()
        {
            return Description;
        }
    }

    public static string ConfigFile(string sShortcut)
    {
        string sConfigFile = "";
        WshShell objShell = new WshShell();
        System.IO.FileInfo fiFile;

        try
        {
            fiFile = new System.IO.FileInfo(sShortcut);
            if (fiFile.Exists == false)
                throw new ApplicationException("Cannot find shortcut " + fiFile.FullName);
            fiFile = new System.IO.FileInfo(sConfigFile);
            if (fiFile.Exists == false)
                throw new ApplicationException("Cannot find config file " + fiFile.FullName);
        }
        catch (Exception ex)
        {
            throw new ApplicationException("General config file error" + "\r\n" + "\r\n" + sShortcut + "\r\n" + "\r\n" + ex.Message);
        }

        return sConfigFile;
    }

    public static string FmtXML(string sString)
    {
        sString = sString.Replace("&amp;", "&");
        sString = sString.Replace("&lt;", "<");
        sString = sString.Replace("&gt;", ">");
        sString = sString.Replace("&quot;", "\"");
        sString = sString.Replace("&apos;", "'");
        return sString.Trim();
    }

    public static string GetSingleNode_Str(XmlNode oXML, string sXPath = "")
    {
        string sReturn = "";

        try
        {
            if (sXPath == "")
                sReturn = FmtXML(oXML.InnerText).Trim();
            else if (oXML.SelectSingleNode(sXPath) == null)
                sReturn = "";
            else
                sReturn = FmtXML(oXML.SelectSingleNode(sXPath).InnerText).Trim();
        }
        catch (Exception ex)
        {
        }

        return sReturn;
    }

    public static int GetSingleNode_Int(XmlNode oXML, string sXPath = "")
    {
        int iReturn = 0;

        try
        {
            if (sXPath == "")
            {
                if (!int.TryParse(FmtXML(oXML.InnerText).Trim(), out iReturn))
                    iReturn = 0;
            }
            else if (oXML.SelectSingleNode(sXPath) == null)
                iReturn = 0;
            else if (!int.TryParse(FmtXML(oXML.SelectSingleNode(sXPath).InnerText).Trim(), out iReturn))
                iReturn = 0;
        }
        catch (Exception ex)
        {
        }

        return iReturn;
    }

    public static decimal GetSingleNode_Dec(XmlNode oXML, string sXPath = "")
    {
        decimal dReturn = 0;

        try
        {
            if (sXPath == "")
            {
                if (!decimal.TryParse(FmtXML(oXML.InnerText).Trim(), out dReturn))
                    dReturn = 0;
            }
            else if (oXML.SelectSingleNode(sXPath) == null)
                dReturn = 0;
            else if (!decimal.TryParse(FmtXML(oXML.SelectSingleNode(sXPath).InnerText).Trim(), out dReturn))
                dReturn = 0;
        }
        catch (Exception ex)
        {
        }

        return dReturn;
    }

    public static bool GetSingleNode_Boo(XmlNode oXML, string sXPath = "")
    {
        bool bReturn = false;
        string sValue = "";

        try
        {
            if (sXPath == "")
                sValue = FmtXML(oXML.InnerText).Trim().ToUpper();
            else if (oXML.SelectSingleNode(sXPath) == null)
                bReturn = false;
            else
                sValue = FmtXML(oXML.SelectSingleNode(sXPath).InnerText).Trim().ToUpper();

            if (sValue == "TRUE")
                bReturn = true;
            else if (sValue == "1")
                bReturn = true;
            else if (sValue == "Y")
                bReturn = true;
        }
        catch (Exception ex)
        {
        }

        return bReturn;
    }

    public static DateTime GetSingleNode_Dte(XmlNode oXML, string sXPath = "")
    {
        try
        {
            if (sXPath == "")
                return DateTime.Parse(oXML.InnerText);
            else if (oXML.SelectSingleNode(sXPath) != null)
                return DateTime.Parse(oXML.SelectSingleNode(sXPath).InnerText);
            else
                return default(DateTime);
        }
        catch (Exception ex)
        {
            return default(DateTime);
        }
    }

    public static DateTime GetSingleNode_DteTime(XmlNode oXML, string sXPath = "")
    {
        try
        {
            if (sXPath == "")
                return DateTime.Parse(oXML.InnerText);
            else if (oXML.SelectSingleNode(sXPath) != null)
                return DateTime.Parse(oXML.SelectSingleNode(sXPath).InnerText);
            else
                return default(DateTime);
        }
        catch (Exception ex)
        {
            return default(DateTime);
        }
    }

    public static string GetSingleNode_Str(XmlDocument oXML, string sXPath = "")
    {
        string sReturn = "";

        try
        {
            if (sXPath == "")
                sReturn = FmtXML(oXML.InnerText).Trim();
            else if (oXML.SelectSingleNode(sXPath) == null)
                sReturn = "";
            else
                sReturn = FmtXML(oXML.SelectSingleNode(sXPath).InnerText).Trim();
        }
        catch (Exception ex)
        {
        }

        return sReturn;
    }

    public static string GetSingleNode_Str_WithERR(XmlDocument oXML, string sXPath = "")
    {
        string sReturn = "";

        try
        {
            if (sXPath == "")
                sReturn = FmtXML(oXML.InnerText).Trim();
            else if (oXML.SelectSingleNode(sXPath) == null)
                throw new Exception("XPATH error");
            else
                sReturn = FmtXML(oXML.SelectSingleNode(sXPath).InnerText).Trim();
        }
        catch (Exception ex)
        {
            throw new Exception("GetSingleNode_Str" + "\n" + "XPATH = " + sXPath + "\n" + ex.Message);
        }

        return sReturn;
    }




    public static int GetSingleNode_Int(XmlDocument oXML, string sXPath = "")
    {
        int iReturn = 0;

        try
        {
            if (sXPath == "")
                iReturn = System.Convert.ToInt32(FmtXML(oXML.InnerText).Trim());
            else if (oXML.SelectSingleNode(sXPath) == null)
                iReturn = 0;
            else
                iReturn = System.Convert.ToInt32(FmtXML(oXML.SelectSingleNode(sXPath).InnerText).Trim());
        }
        catch (Exception ex)
        {
        }

        return iReturn;
    }

    public static decimal GetSingleNode_Dec(XmlDocument oXML, string sXPath = "")
    {
        decimal dReturn = 0;

        try
        {
            if (sXPath == "")
                dReturn = System.Convert.ToDecimal(FmtXML(oXML.InnerText).Trim());
            else if (oXML.SelectSingleNode(sXPath) == null)
                dReturn = 0;
            else
                dReturn = System.Convert.ToDecimal(FmtXML(oXML.SelectSingleNode(sXPath).InnerText).Trim());
        }
        catch (Exception ex)
        {
        }

        return dReturn;
    }

    public static bool GetSingleNode_Boo(XmlDocument oXML, string sXPath = "")
    {
        bool bReturn = false;
        string sValue = "";

        try
        {
            if (sXPath == "")
                sValue = FmtXML(oXML.InnerText).Trim().ToUpper();
            else if (oXML.SelectSingleNode(sXPath) == null)
                bReturn = false;
            else
                sValue = FmtXML(oXML.SelectSingleNode(sXPath).InnerText).ToString().ToUpper();

            if (sValue == "TRUE")
                bReturn = true;
            else if (sValue == "1")
                bReturn = true;
            else if (sValue == "Y")
                bReturn = true;
        }
        catch (Exception ex)
        {
        }

        return bReturn;
    }

    public static DateTime GetSingleNode_Dte(XmlDocument oXML, string sXPath = "")
    {
        DateTime dReturn = default(DateTime);
        if (sXPath == "")
            DateTime.TryParse(oXML.InnerText, out dReturn);
        else if (oXML.SelectSingleNode(sXPath) != null)
            DateTime.TryParse(oXML.SelectSingleNode(sXPath).InnerText, out dReturn);
        return dReturn;
    }

    public static DateTime GetSingleNode_DteTime(XmlDocument oXML, string sXPath = "")
    {
        DateTime dReturn = default(DateTime);
        if (sXPath == "")
            DateTime.TryParse(oXML.InnerText, out dReturn);
        else if (oXML.SelectSingleNode(sXPath) != null)
            DateTime.TryParse(oXML.SelectSingleNode(sXPath).InnerText, out dReturn);
        return dReturn;
    }

    public static string FmtDirectoryName(string sString)
    {
        sString = sString.Trim();
        if (sString != "")
        {
            if (sString.Substring(sString.Length, 0) != @"\")
                sString = sString + @"\";
        }
        return sString;
    }

    public static bool BuildDirectoryTree(string sDirectory)
    {
        string[] arrDirectory;
        string sWrkDirectory = "";
        int i;
        bool bStarted = false;
        bool bfirst = true;
        bool bRC = false;

        try
        {
            arrDirectory = sDirectory.Split('\\');
            for (i = 0; i <= (arrDirectory.Length - 1); i++)
            {
                if ((!bStarted) & (arrDirectory[i]).Trim() != "")
                    bStarted = true;
                if (bStarted)
                {
                    if (bfirst)
                    {
                        if (arrDirectory[i].Substring(2, 1) == ":")
                            sWrkDirectory = arrDirectory[i] + @"\";
                        else
                            sWrkDirectory = @"\\" + arrDirectory[i] + @"\";
                        bfirst = false;
                    }
                    else if (arrDirectory[i].Trim() != "")
                    {
                        sWrkDirectory = sWrkDirectory + arrDirectory[i] + @"\";
                        if (System.IO.Directory.Exists(sWrkDirectory) == false)
                            System.IO.Directory.CreateDirectory(sWrkDirectory);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            bRC = true;
        }
        return bRC;
    }

    public static bool FileIsLocked(System.IO.FileInfo FileFI)
    {
        bool sFileIsLocked = false;
        int iCT;
        bool bOK;

        bOK = false;
        iCT = 0;

        if (FileFI.Exists == false)
            return false;

        while (!bOK)
        {
            try
            {
                if (System.IO.File.Exists(FileFI.FullName + ".LOCK"))
                    System.IO.File.Delete(FileFI.FullName + ".LOCK");
            }
            catch (Exception ex)
            {
            }
            try
            {
                System.IO.File.Move(FileFI.FullName, FileFI.FullName + ".LOCK");
            }
            catch (Exception ex)
            {
            }

            try
            {
                if (System.IO.File.Exists(FileFI.FullName + ".LOCK"))
                {
                    System.IO.File.Move(FileFI.FullName + ".LOCK", FileFI.FullName);
                    sFileIsLocked = false;
                    bOK = true;
                }
                else
                {
                    Thread.Sleep(5000);
                    iCT += 1;
                    if (iCT >= 9)
                    {
                        sFileIsLocked = true;
                        bOK = true;
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        return sFileIsLocked;
    }

    internal static string LogFileName(string sLogName, string sLogDirectory)
    {
        string sDay;
        string sMonth;
        string sYear;
        DateTime dteToday = DateTime.Today;

        string sFileName = "";

        sDay = DateTime.Today.ToString("dd");
        sMonth = DateTime.Today.ToString("MM");
        sYear = DateTime.Today.ToString("yyyy");
        sFileName = sLogDirectory + sLogName + "_" + sYear + sMonth + sDay + ".log";

        return sFileName;
    }

    internal static void CreateLogFile(string sLogName, string sLogDirectory)
    {
        string sLogFile;

        StreamWriter sWriter;
        string sDay;
        string sMonth;
        string sYear;

        try
        {
            sDay = DateTime.Today.ToString("dd");
            sMonth = DateTime.Today.ToString("MMM");
            sYear = DateTime.Today.ToString("yyyy");

            BuildDirectoryTree(sLogDirectory);
            sLogFile = LogFileName(sLogName, sLogDirectory);

            if (System.IO.File.Exists(sLogFile) == false)
            {
                sWriter = System.IO.File.CreateText(sLogFile);
                sWriter.WriteLine("Date: " + sDay + "-" + sMonth + "-" + sYear + "\r\n" + "\r\n");
                sWriter.WriteLine("LOGGING INFORMATION" + "\r\n" + "===================" + "\r\n");
                sWriter.Close();
            }
        }
        catch (Exception ex)
        {
            throw new ApplicationException("Error encounterd whilst greating Logfile  " + sLogName + " in directory" + "\r\n" + sLogDirectory + "\r\n" + ex.Message);
        }
    }

    internal static void LogMessage(string sLogName, string sLogDirectory, string sMessageText)
    {
        System.IO.StreamWriter sWriter;
        string sLogFile;
        int iOK = 0;

        while (iOK != 1)
        {
            try
            {
                sLogFile = LogFileName(sLogName, sLogDirectory);
                sWriter = new System.IO.StreamWriter(sLogFile, true); // NOTE: The "True" parameter causes the text to append, rather than replace
                sWriter.WriteLine(DateTime.UtcNow.ToLocalTime().ToString("HH:mm:ss") + "\r\n" + "~~~~~~~~" + "\r\n" + sMessageText + "\r\n");
                sWriter.Close();
                iOK = 1;
            }
            catch
            {
                Thread.Sleep(2000);
            }
        }
    }

    internal static void SendMail(string sFrom, string EmailFrom, string EmailAddress, string CCAddress, string BCCAddress, string sSubject, string sBody, string SMTPServer, List<Attachment> attachments)
    {
        MailMessage oMail = new MailMessage();

        oMail.From = new MailAddress(sFrom, EmailFrom);
        oMail.To.Add(EmailAddress);
        if (CCAddress.Contains("@"))
        {
            oMail.CC.Add(CCAddress);
        }
        if (BCCAddress.Contains("@"))
        {
            oMail.Bcc.Add(BCCAddress);
        }
        oMail.Subject = sSubject;
        oMail.Body = sBody;
        oMail.IsBodyHtml = true;
        if (attachments != null)
        {
            foreach (Attachment attachment in attachments)
            {
                oMail.Attachments.Add(attachment);
            }
        }
        System.Net.Mail.AlternateView htmlView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(sBody, null, "text/html");

        oMail.AlternateViews.Add(htmlView);

        TrySend(oMail, SMTPServer);

        oMail = null;
    }

    internal static bool TrySend(System.Net.Mail.MailMessage oMail, string sSMTP)
    {
        System.Net.Mail.SmtpClient oSMTP = new System.Net.Mail.SmtpClient(sSMTP, 25);

        try
        {
            oSMTP.Send(oMail);
            return true;
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
            return false;
        }
    }

    internal static void TolerantMove(string sSourceFile, string sTargetFolder)
    {
        FileInfo fiSourceFile;
        FileInfo fiTargetFile;
        int iCT;
        bool bOK;

        // Move the file to the Target Folder
        fiSourceFile = new FileInfo(sSourceFile);

        bOK = false;
        iCT = 0;
        while (!bOK)
        {
            fiTargetFile = new FileInfo(sTargetFolder + fiSourceFile.Name);
            try
            {
                if (fiSourceFile.Exists)
                {
                    if (fiTargetFile.Exists)
                        fiTargetFile.Delete();
                    fiSourceFile.MoveTo(fiTargetFile.FullName);
                }
                bOK = true;
            }
            catch (Exception ex)
            {
                iCT += 1;
                if (iCT >= 9)
                {
                    throw new ApplicationException("Unable to move file " + fiSourceFile.Name);
                    return;
                }
            }
        }
    }
    internal static void TolerantMoveNewName(string sSourceFile, string sTargetFolder, string sNewFilename)
    {
        FileInfo fiSourceFile;
        FileInfo fiTargetFile;
        int iCT;
        bool bOK;

        // Move the file to the Target Folder
        fiSourceFile = new FileInfo(sSourceFile);

        bOK = false;
        iCT = 0;
        while (!bOK)
        {
            fiTargetFile = new FileInfo(sTargetFolder + sNewFilename);
            try
            {
                if (fiSourceFile.Exists)
                {
                    if (fiTargetFile.Exists)
                        fiTargetFile.Delete();
                    fiSourceFile.MoveTo(fiTargetFile.FullName);
                }
                bOK = true;
            }
            catch (Exception ex)
            {
                iCT += 1;
                if (iCT >= 9)
                {
                    throw new ApplicationException("Unable to move file " + fiSourceFile.Name);
                    return;
                }
            }
        }
    }

    internal static void TolerantDelete(string sSourceFile)
    {
        FileInfo fiSourceFile;
        int iCT;
        bool bOK;

        // Move the file to the Target Folder
        fiSourceFile = new FileInfo(sSourceFile);

        bOK = false;
        iCT = 0;
        while (!bOK)
        {
            try
            {
                if (fiSourceFile.Exists)
                    fiSourceFile.Delete();
                bOK = true;
            }
            catch (Exception ex)
            {
                iCT += 1;
                if (iCT >= 9)
                {
                    throw new ApplicationException("Unable to move file " + fiSourceFile.Name);
                    return;
                }
            }
        }
    }

    public static DateTime StringToDate(string sDate)
    {
        DateTime dtReturn;

        dtReturn = (DateTime)TypeDescriptor.GetConverter(new DateTime(1990, 5, 6)).ConvertFrom(sDate);

        return dtReturn;
    }


    public static bool ValidTime(string sTime)
    {

        // Expects time in the form "14:32"

        bool bReturn = false;
        string[] sParts;
        int iHour;
        int iMinute;
        DateTime dtTest;

        try
        {
            sParts = sTime.Split(':');
            iHour = System.Convert.ToInt32(sParts[0]);
            iMinute = System.Convert.ToInt32(sParts[1]);

            if (iHour < 0)
                return false;
            if (iHour > 23)
                return false;
            if (iMinute < 0)
                return false;
            if (iMinute > 59)
                return false;

            dtTest = new DateTime(2000, 1, 1, iHour, iMinute, 0);

            bReturn = true;
        }
        catch (Exception ex)
        {
        }

        return bReturn;
    }

    //public static bool InTimeRange(string sStartTime, string sEndTime)
    //{

    //    // Expects times in the form "14:32"

    //    bool bReturn = false;

    //    DateTime dtNow;
    //    DateTime dtStart;
    //    DateTime dtEnd;
    //    string[] sParts;
    //    int iStartHour;
    //    int iStartMinute;
    //    int iEndHour;
    //    int iEndMinute;

    //    try
    //    {
    //        sParts = sStartTime.Split(new[] { ":" }, StringSplitOptions.None);
    //        iStartHour = System.Convert.ToInt32(sParts[0]);
    //        iStartMinute = System.Convert.ToInt32(sParts[1]);

    //        sParts = sEndTime.Split(new[] { ":" }, StringSplitOptions.None);
    //        iEndHour = System.Convert.ToInt32(sParts[0]);
    //        iEndMinute = System.Convert.ToInt32(sParts[1]);

    //        dtStart = new DateTime(2000, 1, 1, iStartHour, iStartMinute, 0);
    //        dtEnd = new DateTime(2000, 1, 1, iEndHour, iEndMinute, 0);

    //        if (dtEnd < dtStart)
    //            dtEnd = dtEnd.AddHours(24);

    //        dtNow = new DateTime(2000, 1, 1, DateTime.Now.Hour, DateTime.Now.Minute, 0);
    //        if (dtNow < dtStart)
    //            dtNow = dtNow.AddHours(24);

    //        if (dtNow >= dtStart & dtNow <= dtEnd)
    //            bReturn = true;
    //    }
    //    catch (Exception ex)
    //    {
    //    }

    //    return bReturn;
    //}

    public static bool InTimeRange(DateTime dtStartTime, DateTime dtEndTime)
    {
        bool bReturn = false;

        DateTime dtNow;

        try
        {
            if (dtEndTime < dtStartTime)
                dtEndTime = dtEndTime.AddHours(24);
            dtNow = DateTime.Now;
            if (dtNow < dtStartTime)
                dtNow = dtNow.AddHours(24);

            if (dtNow >= dtStartTime & dtNow <= dtEndTime)
                bReturn = true;
        }
        catch (Exception ex)
        {
        }

        return bReturn;
    }

}
