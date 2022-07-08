using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using LMSBot.DAL;
using System.Data.Common;
using System.Collections;
using System.IO;

namespace LMSBot
{
    public class ProcessPayPeriods
    {

        private ArrayList jobs;
        private string _fileDirectoryPath;
        private string JobDate;
        public ProcessPayPeriods(string date)
        {           
            JobDate = date;
            Utility.LogMessage("LMSBot Utility initilized for " + JobDate);
            jobs = GetJobs(JobDate);
            _fileDirectoryPath = Utility.FileDirectory();
        }

        public int ProcessJobs()
        {
            int processedjobs = 0;
            string[] datevals = JobDate.Split('/');

            string filename = string.Format("BalanceFileGenerationSummary{0}{1}{2}.txt",
                                datevals[2], datevals[0].PadLeft(2, '0'), datevals[1].PadLeft(2, '0'));

            string summaryfile = Path.Combine(_fileDirectoryPath, filename);

            FileInfo check = new FileInfo(summaryfile);

            if (check.Exists)
            {
                Utility.LogMessage(string.Format("File '{0}' already exists in '{1}' directory.", filename, _fileDirectoryPath));
                check.Delete();
                Utility.LogMessage(string.Format("Deleted '{0}' from '{1}' directory.", filename, _fileDirectoryPath));
            }
            check = null;

            StreamWriter sw = new StreamWriter(summaryfile, true);
            sw.WriteLine("[" + DateTime.Now.ToString() + "] " + jobs.Count + " file(s) to be generated.");

            foreach (Utility.Job j in jobs)
            {
                string message = string.Empty;

                bool created = GenerateBalanceFile(j, JobDate, out message);

                sw.WriteLine("[" + DateTime.Now.ToString() + "] " + message);

                if (created)
                    processedjobs++;
            }

            sw.WriteLine("[" + DateTime.Now.ToString() + "] " + string.Format("{0} file(s) generated and emailed.", processedjobs));

            sw.Flush();
            sw.Close();
            sw.Dispose();

            Utility.SendEmail("BalanceSummary", "", summaryfile, JobDate, "", "");

            return processedjobs;
        }

        public int ProcessJob(int companyid, int departmentid)
        {
            _fileDirectoryPath = Utility.WebFileDirectory();

            int processedjobs = 0;
            string[] datevals = JobDate.Split('/');

            string filename = string.Format("BalanceFileGenerationSummaryRegen{0}{1}{2}.txt",
                                datevals[2], datevals[0].PadLeft(2, '0'), datevals[1].PadLeft(2, '0'));

            string summaryfile = Path.Combine(_fileDirectoryPath, filename);

            FileInfo check = new FileInfo(summaryfile);

            if (check.Exists)
            {
                Utility.LogMessage(string.Format("File '{0}' already exists in '{1}' directory.", filename, _fileDirectoryPath));
                check.Delete();
                Utility.LogMessage(string.Format("Deleted '{0}' from '{1}' directory.", filename, _fileDirectoryPath));
            }
            check = null;

            StreamWriter sw = new StreamWriter(summaryfile, true);
            sw.WriteLine("[" + DateTime.Now.ToString() + "] " + jobs.Count + " file(s) to be generated.");

            foreach (Utility.Job j in jobs)
            {
                string message = string.Empty;
                bool created = false;
                if (j.CompanyId == companyid && j.DepartmentId == departmentid)
                {
                    created = GenerateBalanceFile(j, JobDate, out message);
                    sw.WriteLine("[" + DateTime.Now.ToString() + "] " + message);
                }

                if (created)
                    processedjobs++;
            }

            sw.WriteLine("[" + DateTime.Now.ToString() + "] " + string.Format("{0} file(s) generated and emailed.", processedjobs));

            sw.Flush();
            sw.Close();
            sw.Dispose();

            Utility.SendEmail("BalanceSummary", "", summaryfile, JobDate, "", "");

            return processedjobs;
        }


       private bool GenerateBalanceFile(Utility.Job _job, string date, out string message)
        {
            bool _completed = false;
            string[] datevals = date.Split('/');
            double cumtotal = 0.0;

            Utility.ExportFileFormat? filefromatinformation = GetFileFormat(_job.CompanyId);

            if (!filefromatinformation.HasValue)
            {
                message = string.Format("No approved file format was found for {0} no file was generated", _job.CompanyName);
                return false;
            }
            Utility.ExportFileFormat fileformat = (Utility.ExportFileFormat)filefromatinformation;

            DataTable filedata = GetFileData(_job.CompanyId, date, _job.SalaryType);

            if (filedata == null)
            {
                message = string.Format("No balance file data was found for the {0} group of {1}, no file was generated", _job.getSalaryType, _job.CompanyName);
                return false;
            }
            if (filedata.Rows.Count == 0)
            {
                message = string.Format("No balance file data was found for the {0} group of {1}, no file was generated", _job.getSalaryType, _job.CompanyName);
                return false;
            }

            string filename = String.Format("{0}-{1}-{2}{3}{4}.csv", _job.CompanyName.Replace(' ', '_'), _job.getSalaryType.Replace(' ', '_'),
                                            datevals[2], datevals[0].PadLeft(2, '0'), datevals[1].PadLeft(2, '0'));
            string balanceFilePath = Path.Combine(_fileDirectoryPath, filename);

            FileInfo check = new FileInfo(balanceFilePath);

            if (check.Exists)
            {
                Utility.LogMessage(string.Format("File '{0}' already exists in '{1}' directory.", filename, _fileDirectoryPath));
                check.Delete();
                Utility.LogMessage(string.Format("Deleted '{0}' from '{1}' directory.", filename, _fileDirectoryPath));
            }
            check = null;

            StreamWriter sw = new StreamWriter(balanceFilePath, true);
            Utility.LogMessage(string.Format("Create '{0}' from '{1}' directory.", filename, _fileDirectoryPath));

            sw.WriteLine(fileformat.ColumnHeadings);

            for (int x = 0; x < filedata.Rows.Count; x++)
            {
                string record = string.Empty;
                DataRow dr = filedata.Rows[x];
                foreach (string dbcolname in fileformat.DBColumnNames)
                {
                    record += dr[dbcolname] + ",";
                }
                record = record.TrimEnd(",".ToCharArray());

                cumtotal += Convert.ToDouble(dr["DeductionAmount"].ToString());

                sw.WriteLine(record);
            }
            sw.Close();

            sw.Dispose();

            Utility.LogMessage(string.Format("Balance file '{0}' with total {1} generated", filename, cumtotal));

            if (_job.CompanyEmail != "")
            {
                bool emailsennt = Utility.SendEmail("CompanyBalanceFile", _job.CompanyEmail, balanceFilePath, _job.CompanyName, _job.getSalaryType, CryptorEngine.Encrypt(_job.PayPeriodId.ToString(), true).Replace("&", "AND").Replace("+", "PLUS").Replace(" ", "SPACE"));
                if (emailsennt)
                {
                    message = string.Format("Balance file with total {2} emailed for the {0} group of {1}", _job.getSalaryType, _job.CompanyName, cumtotal);
                    UpdateFileSent(_job.CompanyId, _job.SalaryType);
                    return true;
                }
            }
            else
            {
                Utility.LogMessage(String.Format("No email address for {0}", _job.CompanyName));
                message = string.Format("Balance file was not emailed for the {0} group of {1}, there was no email address in the database", _job.getSalaryType, _job.CompanyName);
                return false;
            }

            message = "";
            return _completed;
        }

        /// <summary>
        /// Gets the export file format
        /// </summary>
        /// <param name="companyId"></param>
        /// <returns></returns>
        private Utility.ExportFileFormat? GetFileFormat(int companyId)
        {
            Utility.ExportFileFormat? eff = new Utility.ExportFileFormat?(new Utility.ExportFileFormat());

            string errormessage = string.Empty;
            DbCommand cmd = DataAccess.GetCommand();
            DbParameter param = DataAccess.GetParameter();
            param.ParameterName = "@companyId";
            param.Value = companyId;
            cmd.CommandText = "usp_lmsbot_GetExportFileFormats";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add(param);

            DataTable results = DataAccess.GetTable(ref cmd, "TS", out errormessage);

            if (!errormessage.Equals(string.Empty))
                throw new Exception(errormessage);

            if (results == null)
            {
                Utility.LogMessage(string.Format("No approved fileformat found for company with id {0}", companyId));
                return null;
            }
            if (results.Rows.Count == 0)
            {
                Utility.LogMessage(string.Format("No approved fileformat found for company with id {0}", companyId));
                return null;
            }

            string cheadings = string.Empty;

            ArrayList dbcolumns = new ArrayList();

            for (int x = 0; x < results.Rows.Count; x++)
            {
                DataRow dr = results.Rows[x];
                dbcolumns.Add(Convert.ToString(dr["dbColumnName"]));
                cheadings += Convert.ToString(dr["columnName"]) + ",";
            }
            cheadings = cheadings.Substring(0, cheadings.Length - 1);

            eff = new Utility.ExportFileFormat(cheadings, dbcolumns);


            Utility.LogMessage(string.Format("Retrieved file format: {0} and db mapping {2} for company id {1}",
                                cheadings, companyId, string.Join(",", (string[])dbcolumns.ToArray(typeof(string)))));

            return eff;
        }

        /// <summary>
        /// Gets an ArrayList of all jobs (due payperoids) for teh date passed.
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        private ArrayList GetJobs(string date)
        {

            string errormessage = string.Empty;
            DbCommand cmd = DataAccess.GetCommand();
            DbParameter param = DataAccess.GetParameter();
            param.ParameterName = "@DATE";
            param.Value = date;
            cmd.CommandText = "usp_lmsbot_GetJobs";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add(param);

            DataTable results = DataAccess.GetTable(ref cmd, "TS", out errormessage);

            if (!errormessage.Equals(string.Empty))
                throw new Exception(errormessage);

            ArrayList ret = Utility.ExtractJobs(results);
            Utility.LogMessage(ret.Count + " job(s) found and extracted.");
            return ret;
        }

        /// <summary>
        /// Gets the balance file data
        /// </summary>
        /// <param name="departmentId">department</param>
        /// <returns>DataTable with records retrieved</returns>
        private DataTable GetFileData(int departmentId, string date)
        {
            string errormessage = string.Empty;
            DbCommand cmd = DataAccess.GetCommand();
            DbParameter[] param = DataAccess.GetParameters(2);
            param[0].ParameterName = "@DepartmentID";
            param[0].Value = departmentId;
            param[1].ParameterName = "@ProcessDate";
            param[1].Value = date;
            cmd.CommandText = "usp_lmsbot_GetBalanceFileData";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddRange(param);

            DataTable results = DataAccess.GetTable(ref cmd, "TS", out errormessage);

            if (!errormessage.Equals(string.Empty))
                throw new Exception(errormessage);

            if (results != null)
                Utility.LogMessage(results.Rows.Count + " emplyee(s) found for department with id " + departmentId);

            return results;
        }


        /// <summary>
        /// Gets the balance file data
        /// </summary>
        /// <param name="departmentId">department</param>
        /// <returns>DataTable with records retrieved</returns>
        private DataTable GetFileData(int companyid, string date, string payrolltype)
        {
            string errormessage = string.Empty;
            DbCommand cmd = DataAccess.GetCommand();
            DbParameter[] param = DataAccess.GetParameters(3);
            param[0].ParameterName = "@CompanyID";
            param[0].Value = companyid;
            param[1].ParameterName = "@ProcessDate";
            param[1].Value = date;
            param[2].ParameterName = "@SalaryType";
            param[2].Value = payrolltype;
            cmd.CommandText = "usp_lmsbot_GetBalanceFileData";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddRange(param);

            DataTable results = DataAccess.GetTable(ref cmd, "TS", out errormessage);

            if (!errormessage.Equals(string.Empty))
            {                
                Utility.LogMessage("Run time error: " + errormessage + " for company with id " + companyid + " payroll group " + payrolltype);
                throw new Exception(errormessage);
            }

            if (results != null)
                Utility.LogMessage(results.Rows.Count + " employee(s) found for company with id " + companyid + " payroll group " + payrolltype);           

            return results;
        }


        /// <summary>
        /// Updates the file sent status of a payperiod to Y
        /// </summary>
        /// <param name="payperiodid">Pay Period record id</param>
        private void UpdateFileSent(int payperiodid)
        {
            string errormessage = string.Empty;
            DbCommand cmd = DataAccess.GetCommand();
            DbParameter[] param = DataAccess.GetParameters(2);
            param[0].ParameterName = "@PayPeriodId";
            param[0].Value = payperiodid;
            param[1].ParameterName = "@MailSentDate";
            param[1].Value = DateTime.Now.ToString("MM/dd/yyyy");
           

            cmd.CommandText = "usp_lmsbot_UpdateFileSentStatus";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddRange(param);

            DataAccess.ExecuteNonQuery(ref cmd, "TS", out errormessage);

            if (!errormessage.Equals(string.Empty))
                throw new Exception(errormessage);
        }


        private void UpdateFileSent(int companyid,string payrolltype)
        {
            string errormessage = string.Empty;
            DbCommand cmd = DataAccess.GetCommand();
            DbParameter[] param = DataAccess.GetParameters(3);
            param[0].ParameterName = "@CompanyID";
            param[0].Value = companyid;
            param[1].ParameterName = "@SalaryType";
            param[1].Value = payrolltype;
            param[2].ParameterName = "@MailSentDate";
            //param[2].Value = DateTime.Now.ToString("MM/dd/yyyy");
            param[2].Value = DateTime.Now.Year.ToString() + "" + DateTime.Now.Month.ToString("0#") + "" + DateTime.Now.Day.ToString("0#");

            cmd.CommandText = "usp_lmsbot_UpdateFileSentStatus";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddRange(param);

            DataAccess.ExecuteNonQuery(ref cmd, "TS", out errormessage);

            if (!errormessage.Equals(string.Empty))
                throw new Exception(errormessage);
        }



    }
}
