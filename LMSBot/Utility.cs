using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Reflection;
using System.Net.Mail;
using System.Net;
using System.Data;
using System.Collections;
using System.Configuration;
using LMSBot.Properties;
using LMSBot.DAL;
using System.Data.Common;


namespace LMSBot
{
    public class Utility
    {

        /// <summary>
        /// Cannot Instantiate this class.
        /// </summary>
        private Utility()
        {

        }

        /// <summary>
        /// Struct to hold informarion about each cutoff date that need to be processed
        /// </summary>
        public struct Job
        {
            /// <summary>
            /// Constructor for Job struct
            /// </summary>
            /// <param name="_ppid">PayPeriodId</param>
            /// <param name="_codate">CutOffDate</param>
            /// <param name="_pdate">PayDate</param>
            /// <param name="_did">DepartmentId</param>
            /// <param name="_dname">Department name</param>
            /// <param name="_cname">Company name</param>
            /// <param name="_cid">CompanyId</param>
            /// <param name="_cid">CompanyEmail</param>
            public Job(int _ppid, DateTime _codate, DateTime _pdate, int _did, string _dname, string _cname, int _cid, string _email, string _DeptSalaryType)
            {
                _PayPeriodId = _ppid;
                _CutOffDate = _codate;
                _PayDate = _pdate;
                _DepartmentId = _did;
                _DepartmentName = _dname;
                _CompanyId = _cid;
                _CompanyName = _cname;
                _CompanyEmail = _email;
                _SalaryType = _DeptSalaryType.ToUpper();

            }

            private int _PayPeriodId;
            public int PayPeriodId
            {
                get
                {
                    return _PayPeriodId;
                }
                set
                {
                    _PayPeriodId = value;
                }
            }

            private DateTime _CutOffDate;
            public DateTime CutOffDate
            {
                get
                {
                    return _CutOffDate;
                }
                set
                {
                    _CutOffDate = value;
                }
            }

            private DateTime _PayDate;
            public DateTime PayDate
            {
                get
                {
                    return _PayDate;
                }
                set
                {
                    _PayDate = value;
                }
            }

            private int _DepartmentId;
            public int DepartmentId
            {
                get
                {
                    return _DepartmentId;
                }
                set
                {
                    _DepartmentId = value;
                }
            }

            private string _DepartmentName;
            public string DepartmentName
            {
                get
                {
                    return _DepartmentName;
                }
                set
                {
                    _DepartmentName = value;
                }
            }

            private int _CompanyId;
            public int CompanyId
            {
                get
                {
                    return _CompanyId;
                }
                set
                {
                    _CompanyId = value;
                }
            }

            private string _CompanyName;
            public string CompanyName
            {
                get
                {
                    return _CompanyName;
                }
                set
                {
                    _CompanyName = value;
                }
            }

            private string _CompanyEmail;
            public string CompanyEmail
            {
                get
                {
                    return _CompanyEmail;
                }
                set
                {
                    _CompanyEmail = value;
                }
            }

            private string _SalaryType;
            public string SalaryType
            {
                get
                {
                    return _SalaryType;
                }
                set
                {
                    _SalaryType = value;
                }
            }

            public string getSalaryType
            {
                get
                {
                    if (_SalaryType == "W")
                    {
                        return "Weekly";
                    }
                    else if (_SalaryType == "M")
                    {
                        return "Monthly";
                    }

                    else if (_SalaryType == "F")
                    {
                        return "Fortnightly";
                    }
                    else {
                        return "_";
                    }
                    
                }

            }

        }

        /// <summary>
        /// 
        /// </summary>
        public struct ExportFileFormat
        {
            string _ColumnHeadings;
            ArrayList _DBColumnNames;

            public ExportFileFormat(string _pcolumnheadings, ArrayList _pdbcolumnnames)
            {
                _ColumnHeadings = _pcolumnheadings;
                _DBColumnNames = _pdbcolumnnames;
            }

            public string ColumnHeadings
            {
                get
                {
                    return _ColumnHeadings;
                }
            }

            public ArrayList DBColumnNames
            {
                get
                {
                    return _DBColumnNames;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dr"></param>
        /// <param name="field"></param>
        /// <returns></returns>
        private static bool CheckNullOrBlank(DataRow dr, string field)
        {
            if (dr[field] != DBNull.Value)
            {
                if (Convert.ToString(dr[field]) != "")
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="_jobs"></param>
        /// <returns></returns>
        public static ArrayList ExtractJobs(DataTable _jobs)
        {
            ArrayList ret = new ArrayList();

            for (int x = 0; x < _jobs.Rows.Count; x++)
            {
                string emailaddress = string.Empty;

                DataRow dr = _jobs.Rows[x];

                if (CheckNullOrBlank(dr, "CompanyEmail"))
                {
                    emailaddress = Convert.ToString(dr["CompanyEmail"]);
                }
                else if (CheckNullOrBlank(dr, "EmailID"))
                {
                    emailaddress = Convert.ToString(dr["CompanyEmail"]);
                }
                else
                {
                    LogMessage(string.Format("No email address found for {0}", Convert.ToString(dr["company"])));
                    emailaddress = "";
                }

                Job j = new Job(Convert.ToInt32(dr["PayPeriodId"]), Convert.ToDateTime(dr["CutOffDate"]), Convert.ToDateTime(dr["PayDate"]),
                                Convert.ToInt32(dr["departmentId"]), Convert.ToString(dr["department"]), Convert.ToString(dr["company"]),
                                Convert.ToInt32(dr["companyId"]), Convert.ToString(dr["CompanyEmail"]), Convert.ToString(dr["SalaryType"]));
                ret.Add(j);
            }

            return ret;
        }

        /// <summary>
        /// Writes a message to the application log file
        /// </summary>
        /// <param name="message">string to be written</param>
        public static void LogMessage(String message)
        {
            if (message == null)
            {
                message = string.Empty;
            }

            String strAssemblyPath = ConfigurationManager.AppSettings["WebFileDirectory"].ToString();
            String logPath = Path.Combine(strAssemblyPath, "Log");

            //Checks if directory exists and create it if it dosn't exists
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }

            String logFilePath = Path.Combine(logPath, String.Format("log{0}{1}{2}.log", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString().PadLeft(2, '0'), DateTime.Now.Day.ToString().PadLeft(2, '0')));

            StreamWriter sw = new StreamWriter(logFilePath, true);

            //Writes line in the format <code>currentdatetime|message</code>
            sw.WriteLine(DateTime.Now.ToString() + "|" + message);
            sw.Flush();
            sw.Close();
        }

        /// <summary>
        /// Writes a line to the application error log
        /// </summary>
        /// <param name="message">Error message</param>
        /// <param name="exceptionMessage">Exception error message</param>
        /// <param name="trace">Exception stack trace</param>
        public static void LogError(string message, string exceptionMessage, string trace)
        {
            if (message == null)
            {
                message = string.Empty;
            }

            if (exceptionMessage == null)
            {
                exceptionMessage = string.Empty;
            }

            if (trace == null)
            {
                trace = string.Empty;
            }
            String strAssemblyPath = ConfigurationManager.AppSettings["WebFileDirectory"].ToString();
            String logPath = Path.Combine(strAssemblyPath, "Log");

            //Checks if directory exists and create it if it dosn't exists
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }

            String logFilePath = Path.Combine(logPath, String.Format("errorlog{0}{1}{2}.log", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString().PadLeft(2, '0'), DateTime.Now.Day.ToString().PadLeft(2, '0')));

            StreamWriter sw = new StreamWriter(logFilePath, true);

            //Writes line in the format <code>currentdatetime|message|exceptionMessage|trace</code>
            sw.WriteLine(DateTime.Now.ToString() + "|" + message + "|" + exceptionMessage + "|" + trace);
            sw.Close();
        }

        /// <summary>
        /// Sends an email
        /// </summary>
        /// <param name="message"></param>
        /// <param name="sub"></param>
        /// <param name="mailId"></param>
        public static bool SendEmail(string emailtype, string mailto, string attachmentPath, string company, string departmentgroup, string payperiodid)
        {
            MailMessage msg = new MailMessage();
            SmtpClient smtpc;
            string fromemailaddress = "";
            string emailbody = string.Empty;
            string emailsubject = string.Empty;
            string ccemail = string.Empty;
            string ccompdept = string.Empty;
            string mailapiuser = string.Empty;
            string mailusername = string.Empty;
            string mailpwd = string.Empty;
            string smtp_address = string.Empty;
            string smtp_port = string.Empty;
            string useSmptSSL = "";
            string[] fromemail;

            try
            {
                fromemailaddress = ConfigurationManager.AppSettings["FROMEMAIL"].ToString();
                if (fromemailaddress.Contains("|"))
                {
                     fromemail = ConfigurationManager.AppSettings["FROMEMAIL"].ToString().Split('|');
                }
                else 
                {
                    fromemailaddress = fromemailaddress + "|" + "ePay Automated Statements";
                    fromemail = fromemailaddress.Split('|');
                }
                string domainurl = ConfigurationManager.AppSettings["DomainRoot"].ToString();
                mailpwd = ConfigurationManager.AppSettings["FROMPWD"].ToString().Trim();
                smtp_address = ConfigurationManager.AppSettings["SMTP"].ToString();
                smtp_port = ConfigurationManager.AppSettings["SMTPPort"];
                mailapiuser = ConfigurationManager.AppSettings["MailAPIUser"].ToString().Trim();
                useSmptSSL= ConfigurationManager.AppSettings["SMTPSSL"].ToString();                                    


                if (emailtype == "CompanyBalanceFile")
                {
                    ccompdept = string.Concat(company, " ", departmentgroup);
                    emailbody = System.Web.HttpUtility.HtmlDecode(Settings.Default.BalanceFileEmailBody).Replace("{0}", company).Replace("{1}", departmentgroup);
                    domainurl += String.Format("CreditPages/ExportFilesStatusReport.aspx?pid={0}", payperiodid);
                    emailbody = emailbody.Replace("{2}", domainurl);
                    emailsubject = Settings.Default.BalanceFileEmailSubject.Replace("{0}", ccompdept);
                    ccemail = ConfigurationManager.AppSettings["CCBalanceEmail"].ToString();
                }

                if (emailtype == "BalanceSummary")
                {
                    emailbody = System.Web.HttpUtility.HtmlDecode(Settings.Default.SummaryFileEmailBody);
                    emailsubject = Settings.Default.SummaryFileEmailSubject.Replace("{0}",company);
                    mailto = ConfigurationManager.AppSettings["CCExecutionReport"].ToString();
                }

                MailAddress fromaddress = new MailAddress(fromemail[0],fromemail[1]);                

                msg.From = fromaddress;
                msg.To.Add(mailto);
                msg.Subject = emailsubject;
                
                if (!string.IsNullOrEmpty(ccemail))
                    msg.CC.Add(ccemail);

                msg.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess;
                msg.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
                msg.Headers.Add("Disposition-Notification-To", "<" + ConfigurationManager.AppSettings["FailedEmailNoticeAddress"].ToString() + ">");
                msg.IsBodyHtml = true;
                msg.Body = emailbody;

                Attachment attachment = new Attachment(attachmentPath); //create the attachment
                msg.Attachments.Add(attachment);
           
                if (!string.IsNullOrEmpty(mailapiuser))
                {
                    mailusername = mailapiuser;                    
                }
                else
                {
                    mailusername = fromemail[0].ToString();                                        
                }
                //smtpc = new SmtpClient(ConfigurationManager.AppSettings["SMTP"].ToString(), Convert.ToInt32(ConfigurationManager.AppSettings["SMTPPort"]));
                //smtpc.Credentials = new NetworkCredential(fromemail[0], ConfigurationManager.AppSettings["FROMPWD"].ToString());
                
                smtpc = new SmtpClient(smtp_address, Convert.ToInt32(smtp_port));                
                smtpc.UseDefaultCredentials = false;
                smtpc.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                smtpc.Credentials = new NetworkCredential(mailusername, mailpwd);                               

                if (useSmptSSL == "true")
                    smtpc.EnableSsl = true;
                else
                    smtpc.EnableSsl = false;

                smtpc.Timeout = 10000;

                smtpc.Send(msg);

                LogMessage("Email sent");

                return true;
            }
            catch (Exception ex)
            {
                LogError("SendEmail() has an error for " + "\r\n" + company + " mailto list: " + mailto + " \r\n", ex.ToString(), ex.StackTrace);
            }
            finally
            {
                msg = null;
                smtpc = null;
            }

            return false;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="company_id"></param>
        /// <returns></returns>
        public static Hashtable GetTransactionCodeMappings(int companyId)
        {
            string errormessage = string.Empty;
            DbCommand cmd = DataAccess.GetCommand();
            DbParameter param = DataAccess.GetParameter();
            param.ParameterName = "@companyId";
            param.Value = companyId;
            cmd.CommandText = "usp_lmsbot_GetTransactionCodes";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add(param);

            DataTable results = DataAccess.GetTable(ref cmd, "TS", out errormessage);
            Hashtable ht = new Hashtable();

            for (int x = 0; x < results.Rows.Count; x++)
            {
                DataRow r = results.Rows[x];
                ht.Add(r["InternalTransactionCode"].ToString(), r["CompanyTransactionCode"].ToString());
            }

            return ht;
        }


        public static string FileDirectory()
        {
            String strAssemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            String path = Path.Combine(strAssemblyPath, "BalanceFiles");
            path = Path.Combine(path, String.Format("{0}{1}{2}", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString().PadLeft(2, '0'), DateTime.Now.Day.ToString().PadLeft(2, '0')));

            //Checks if directory exists and create it if it dosn't exists
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            return path;
        }


        public static string WebFileDirectory()
        {
            String strAssemblyPath = ConfigurationManager.AppSettings["WebFileDirectory"].ToString();
            String path = Path.Combine(strAssemblyPath, "BalanceFiles");
            path = Path.Combine(path, String.Format("{0}{1}{2}", DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString().PadLeft(2, '0'), DateTime.Now.Day.ToString().PadLeft(2, '0')));

            //Checks if directory exists and create it if it dosn't exists
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            return path;
        }
    }
}
