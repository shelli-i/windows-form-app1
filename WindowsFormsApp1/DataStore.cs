using System;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;

namespace WindowsFormsApp1
{
    internal static class DataStore
    {
        /// <summary>
        /// For connection string injection
        /// </summary>
        /// <param name="connString"></param>
        /// <param name="provider"></param>
        /// <returns>DbCommand</returns>
        internal static DbCommand CreateSpCommand(
            string connString,
            string provider)
        {
            var factory = DbProviderFactories.GetFactory(
                provider);

            var conn = factory.CreateConnection();

            conn.ConnectionString = connString;

            var cmd = conn.CreateCommand();

            cmd.CommandType = CommandType.StoredProcedure;

            return cmd;
        }

        /// <summary>
        /// For connection string injection
        /// </summary>
        /// <param name="connString"></param>
        /// <param name="provider"></param>
        /// <returns>DbCommand</returns>
        internal static DbCommand CreateTextCommand(
            string connString,
            string provider)
        {
            var factory = DbProviderFactories.GetFactory(
                provider);

            var conn = factory.CreateConnection();

            conn.ConnectionString = connString;

            var cmd = conn.CreateCommand();

            cmd.CommandType = CommandType.Text;

            return cmd;
        }

        /// <summary>
        /// Returns appropriate DataReader based on provider
        /// </summary>
        /// <param name="cmd"></param>
        /// <returns>DbDataReader</returns>
        internal static DbDataReader GetDbDataReader(
            DbCommand cmd)
        {
            try
            {
                cmd.Connection.Open();

                return cmd.ExecuteReader(
                    CommandBehavior.CloseConnection);
            }
            catch (DbException de)
            {
                throw de;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Executes T-SQL Statement and indicates success or failure
        /// </summary>
        /// <param name="cmd"></param>
        /// <returns>bool</returns>
        internal static bool ExecuteNonQuery(
            DbCommand cmd)
        {
            try
            {
                cmd.Connection.Open();

                return ((cmd.ExecuteNonQuery() != 0) ? true : false);
            }
            catch (DbException de)
            {
                throw de;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (cmd != null)
                    cmd.Connection.Close();
            }
        }

        /// <summary>
        /// Executes a T-SQL statement and returns only a single value
        /// </summary>
        /// <param name="cmd"></param>
        /// <returns>int</returns>
        internal static int ExecuteScalar(
            DbCommand cmd)
        {
            try
            {
                cmd.Connection.Open();

                return Int32.Parse(cmd.ExecuteScalar().ToString());
            }
            catch (DbException de)
            {
                throw de;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (cmd != null)
                    cmd.Connection.Close();
            }
        }
    }

    internal class DbProviderFactories
    {
        internal static DbProviderFactory GetFactory(string providerInvariantName)
        {
            DbProviderFactory factory = null;
            switch (providerInvariantName)
            {
                case "System.Data.SqlClient":
                    factory = SqlClientFactory.Instance;
                    break;
                default:
                    throw new NotImplementedException(
                        "This has only been configured to use SqlServer's native client at this time.  Provider not implemented");
            }
            return factory;
        }
    }
}
