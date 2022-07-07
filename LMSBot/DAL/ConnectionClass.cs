using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.Common;
using System.Configuration;

namespace LMSBot.DAL
{
    public class ConnectionClass   
    {
       
        /// <summary>
        /// Cannot Instantiate this class
        /// </summary>
        private ConnectionClass()
        {
            
        }

        /// <summary>
        /// Returns a closed DbConnection object. User has to open/close it whenever required.
        /// </summary>
        /// <returns></returns>
        public static DbConnection GetConnection(string ConnectionStringName)
        {
            //Dont let the user to create any new connections.
            try
            {
                string ConnectionString = Convert.ToString(ConfigurationManager.ConnectionStrings[ConnectionStringName + "ConnectionString"]);
                DbProviderFactory Dbfactory = ConnectionClass.GetDbFactory();
                DbConnection conn = Dbfactory.CreateConnection();
                conn.ConnectionString = ConnectionString;
                return conn;
            }
            catch (DbException)
            {
                throw new Exception("An exception has occured while creating the connection. Please check Connection String settings in the web.config file.");
            }
        
        }



        /// <summary>
        /// Provides with a DbProviderFactory object with the provider name from the config file.
        /// </summary>
        /// <returns>DbProviderFactory object</returns>
        internal  static DbProviderFactory GetDbFactory()
        {
            try
            {
                string ProviderName = ConfigurationManager.AppSettings["ProviderName"];
                DbProviderFactory Dbfactory = DbProviderFactories.GetFactory(ProviderName);
                return Dbfactory;
            }
            catch(DbException)
            {
                throw new Exception("An exception has occured while creating the database provider factory. Please check the ProviderName specified in the web.config file.");
            }
        }

        
        /// <summary>
        /// Provides with a DbProviderFactory object with the supplied Provider name.
        /// </summary>
        /// <param name="ProviderName">Data Provider Name (e.g.) Oledb, Odbc, SqlClient</param> 
        /// <returns>DbProviderFactory object</returns>
        internal  static DbProviderFactory GetDbFactory(string ProviderName)
        {
            DataTable dtProviders = DbProviderFactories.GetFactoryClasses();

            if (dtProviders.Rows.Count == 0)
            {
                throw new Exception("No Data Providers are installed in the .Net FrameWork that implement the abstract DbProviderFactory Classes. "); 
            }
           
                bool errorFlag = false;
                foreach (DataRow dr in dtProviders.Rows)
                {
                    if (dr[2] != null)
                    {
                        string ExistingProviderName = dr[2].ToString();
                        if (ProviderName.ToLower() == ExistingProviderName.Trim().ToLower())
                        {
                            errorFlag = false;
                            break;
                        }
                        else
                        {
                            errorFlag = true;
                        }

                    }
                }

                if (errorFlag)
                {
                    throw new Exception("The ProviderName string supplied is not a valid Provider Name<BR>or it does not implement the abstract DbProviderFactory Classes. <BR>The string ProviderName is case-sensitive. Also please check it for proper spelling. ");
                }
                DbProviderFactory Dbfactory = DbProviderFactories.GetFactory(ProviderName);
                return Dbfactory;
        }

    }
}
