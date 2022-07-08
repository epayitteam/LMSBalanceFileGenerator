using System;
using System.Collections.Generic;
using System.Text;
using LMSBot;

namespace LMSBalanceFileGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            ProcessPayPeriods ppp;

            try
            {
                //ppp = new ProcessPayPeriods(DateTime.Now.ToString("MM/dd/yyyy"));               
                ppp = new ProcessPayPeriods(DateTime.Now.Month.ToString("0#") + "/" + DateTime.Now.Day.ToString("0#") + "/" + DateTime.Now.Year.ToString("00##"));                                                            
                ppp.ProcessJobs();
                Utility.LogMessage("LMSBot Utility completed");
            }
            catch (Exception ex)
            {
                string innermessage = string.Empty;

                if (ex.InnerException != null)
                    innermessage = ex.InnerException.Message;

                Utility.LogError("Error occurred in main routine: " + innermessage, ex.Message, ex.StackTrace);
            }

        }
    }
}
