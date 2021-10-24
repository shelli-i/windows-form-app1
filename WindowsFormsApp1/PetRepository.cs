using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using WindowsFormsApp1.Model;

namespace WindowsFormsApp1
{
    public class PetRepository : IPetRepository
    {
        private readonly ConnectionStringSettings _conn;
        private readonly PetDataSetFactory _factory;

        public PetRepository(
            ConnectionStringSettings saConn)
            : base()
        {
            var msg = $"{this.GetType().Name} expects ctor injection.";

            this._conn = saConn ?? throw new ArgumentNullException(
                msg);

            this._factory = new PetDataSetFactory();
        }


        public IList<PetData> FindAllFiles()
        {
            DbCommand cmd = null;
            IDataReader dr = null;
            try
            {
                cmd = DataStore.CreateSpCommand(
                    this._conn.ConnectionString,
                    this._conn.ProviderName);

                cmd.CommandText = SqlStoredProcs.PetDataProcs.usp_GetDataSources.ToString();

                var list = new List<PetData>();

                dr = DataStore.GetDbDataReader(cmd);
                while (dr.Read())
                {
                    list.Add(
                        new PetData
                        { 
                            DataSrcID = Int32.Parse(dr["DataSrcID"].ToString()),
                            fileName = dr["File Name"].ToString(),
                            userID = dr["userID"].ToString(),
                            DateEntry = dr["DateEntry"].ToString(),
                            Count = Int32.Parse(dr["Count"].ToString())
                        });
                }

                return list;
            }
            catch (DbException de)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                if (dr != null)
                    dr.Close();
                if (cmd != null)
                    cmd.Connection.Close();
            }
        }
        public void PrintExcelFile(int DataId, string fname)
        {
            SqlCommand cmd = null;
            string result = null;
            try
            {
                SqlConnection con = new SqlConnection(this._conn.ConnectionString);
                cmd = new SqlCommand(SqlStoredProcs.PetDataProcs.usp_GetExcelFile.ToString(), con)
                {
                    CommandTimeout = 600
                };

                var ds = new DataSet();
             
                SqlDataAdapter da = new SqlDataAdapter(cmd);              
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@DataSrcID", SqlDbType.Int).Value = DataId;
                cmd.Connection.Open();
                da.Fill(ds);                  

                result = String.Format(fname, DataId.ToString());

                using (ExcelPackage exPkg = new ExcelPackage())
                {
                        ExcelWorksheet ws = exPkg.Workbook.Worksheets.Add("sheetMain");
                        ws.Cells["A1"].LoadFromDataTable(ds.Tables[0], true);
                        ws.Cells.AutoFitColumns();
                        FileInfo fi = new FileInfo(result);
                        exPkg.SaveAs(fi);
                }
                    Process.Start(@"cmd.exe ", @"/c " + result);
             
               // return result;
            }
            catch (DbException de)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
             
                if (cmd != null)
                    cmd.Connection.Close();
            }
        }
    }
}
