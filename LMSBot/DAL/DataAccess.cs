#region using
using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.Common;
using System.Configuration;
#endregion

namespace LMSBot.DAL
{
    /// <summary>
    /// This Class contains all the Utility methods to interact with database.
    /// </summary>
    public class DataAccess
    {
        /// <summary>
        /// Cannot Instantiate this class.
        /// </summary>
        private DataAccess()
        {

        }

        #region Utility Methods

        /// <summary>
        /// Returns a DataSet filled with data from the execution of the given Command.
        /// </summary>
        /// <param name="command">Command Object filled with necessary parameters.</param>
        /// <param name="ErrorMessage">Output parameter for getting the error message if any.</param>
        /// <returns>Returns a DataSet filled with data from the execution of the given Command.</returns>
        public static DataSet GetDataSet(ref DbCommand command, string DBName, out string ErrorMessage)
        {
            ErrorMessage = String.Empty;
            DataSet ds = new DataSet();
            DbDataAdapter dbDap = null;
            DbConnection conn = null;

            if (command == null)
            {
                ErrorMessage = "Please initilise the command object.";
                return null;
            }

            try
            {
                conn = ConnectionClass.GetConnection(DBName);
                command.Connection = conn;
                conn.Open();
                DbProviderFactory Dbfactory = ConnectionClass.GetDbFactory();
                dbDap = Dbfactory.CreateDataAdapter();
                dbDap.SelectCommand = command;
                dbDap.Fill(ds);
            }
            catch (DbException exp)
            {
                ErrorMessage = "An exception has occured while executing the database transactions.  <BR>";
                ErrorMessage = ErrorMessage + exp.Message;
            }
            finally
            {
                if (command != null)
                {
                    command.Dispose();
                }
                if (dbDap != null)
                {
                    dbDap.Dispose();
                }
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
                if (conn != null)
                {
                    conn.Dispose();
                }
            }

            return ds;

        }



        /// <summary>
        /// Executes the command object returning an int value.
        /// </summary>
        /// <param name="command">Command object filled with the necessary parameters.</param>
        /// <param name="ErrorMessage">Error Message if any.</param>
        /// <returns>Int32 value indicating the error.</returns>
        public static int ExecuteNonQuery(ref DbCommand command, string DBName, out string ErrorMessage)
        {
            ErrorMessage = String.Empty;
            DbTransaction tran = null;
            DbConnection conn = null;
            int result = 0;

            if (command == null)
            {
                ErrorMessage = "Please initilise the command object.";
                return result = -1;
            }

            try
            {
                conn = ConnectionClass.GetConnection(DBName);
                command.Connection = conn;
                conn.Open();
                tran = conn.BeginTransaction();
                command.Transaction = tran;
                result = command.ExecuteNonQuery();
                tran.Commit();
            }
            catch (DbException exp)
            {
                ErrorMessage = "An exception has occured while executing the database transactions.  <BR>";
                ErrorMessage = ErrorMessage + exp.Message;
                if (tran != null)
                {
                    tran.Rollback();
                }
            }

            finally
            {
                if (command != null)
                {
                    command.Dispose();
                }
                if (tran != null)
                {
                    tran.Dispose();
                }
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
                if (conn != null)
                {
                    conn.Dispose();
                }
            }

            return result;
        }



        /// <summary>
        /// Executes the command object returning the first column of the first row.
        /// </summary>
        /// <param name="command">Command object filled with the necessary parameters.</param>
        /// <param name="ErrorMessage">Error Message if any.</param>
        /// <returns>An object containing the first column of the first row.</returns>
        public static object ExecuteScalar(ref DbCommand command, string DBName, out string ErrorMessage)
        {
            ErrorMessage = String.Empty;
            DbTransaction tran = null;
            DbConnection conn = null;
            object result = null;

            if (command == null)
            {
                ErrorMessage = "Please initilise the command object.";
                return null;
            }

            try
            {
                conn = ConnectionClass.GetConnection(DBName);
                command.Connection = conn;
                conn.Open();
                // Begin transaction
                tran = conn.BeginTransaction();
                // Assign connection transaction to command
                command.Transaction = tran;
                //execute
                result = command.ExecuteScalar();
                if (tran != null)
                {
                    //commit if no error
                    tran.Commit();
                }
            }
            catch (DbException exp)
            {
                ErrorMessage = "An exception has occured while executing the database transactions.  <BR>";
                ErrorMessage = ErrorMessage + exp.Message;
                if (tran != null)
                {
                    //roll back if error
                    tran.Rollback();
                }
            }

            finally
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
                if (conn != null)
                {
                    conn.Dispose();
                }
                if (command != null)
                {
                    command.Dispose();
                }
                if (tran != null)
                {
                    tran.Dispose();
                }

            }

            return result;
        }



        /// <summary>
        /// Returns a DbCommand object.
        /// </summary>
        /// <returns>Returns a DbCommand object.</returns>
        public static DbCommand GetCommand()
        {
            DbCommand comm = ConnectionClass.GetDbFactory().CreateCommand();
            return comm;
        }



        /// <summary>
        /// Returns an array of DbParameter objects.
        /// </summary>
        /// <param name="count">Denotes the size of the required parameter array. Should be greater than 0.</param>
        /// <returns>Returns an array of DbParameter objects.</returns>
        public static DbParameter[] GetParameters(ushort count)
        {
            DbProviderFactory Dbfactory = ConnectionClass.GetDbFactory();
            DbParameter[] parameters = new DbParameter[count];

            try
            {
                for (int i = 0; i < count; i++)
                {
                    parameters[i] = Dbfactory.CreateParameter();
                }
            }
            catch (System.IndexOutOfRangeException)
            {
                throw new Exception("Count has to be greater than zero and less than 100", new IndexOutOfRangeException());
            }

            return parameters;
        }



        /// <summary>
        /// Returns a single DbParameter object 
        /// </summary>
        /// <returns>Returns a single DbParameter object </returns>
        public static DbParameter GetParameter()
        {
            DbProviderFactory Dbfactory = ConnectionClass.GetDbFactory();
            DbParameter parameter = Dbfactory.CreateParameter();
            return parameter;
        }



        /// <summary>
        /// Uses a fast datareader to load and return a datatable.
        /// </summary>
        /// <param name="command">Command object filled with necessary parameters.</param>
        /// <param name="ErrorMessage">Output parameter for getting the error message if any.</param>
        /// <returns>Returns a loaded DataTable from the execution of the given Command.</returns>
        public static DataTable GetTable(ref DbCommand command, string DBName, out string ErrorMessage)
        {
            ErrorMessage = String.Empty;
            DbTransaction tran = null;
            DbConnection conn = null;
            DbDataReader reader = null;
            DataTable dtable = new DataTable();

            if (command == null)
            {
                ErrorMessage = "Please initilise the command object.";
                return dtable;
            }

            try
            {
                conn = ConnectionClass.GetConnection(DBName);
                command.Connection = conn;
                conn.Open();
                tran = conn.BeginTransaction();
                command.Transaction = tran;
                reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    dtable.Load(reader);
                }
                else
                {
                    //ErrorMessage = "No records found.  <BR>";
                    return dtable;
                }
            }
            catch (DbException exp)
            {
                ErrorMessage = "An exception has occured while executing the database transactions.  <BR>";
                ErrorMessage = ErrorMessage + exp.Message;
                if (tran != null)
                {
                    tran.Rollback();
                }
            }

            finally
            {
                command.Dispose();
                if (tran != null)
                {
                    if (reader != null)
                        reader.Close();
                    
                    tran.Dispose();
                }
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
                if (conn != null)
                {
                    conn.Dispose();
                }
            }

            return dtable;
        }

        #endregion

    }
}
