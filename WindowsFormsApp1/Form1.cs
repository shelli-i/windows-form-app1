using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using OfficeOpenXml;
using Autofac;
using Microsoft.Extensions.Configuration;
using System.Configuration;
using System.Security.Principal;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private readonly IPetService _petSvc;
        private string _userName;
        public static string FileName = ConfigurationManager.AppSettings["FileName"].ToString();
     //   public static string sqlCn = ConfigurationManager.AppSettings["ConnectionStrings:SQLConn"].ToString();
        public Form1(IPetService petSvc)
        {
            var msg = $"{this.GetType().Name} expects ctor injection.";

            this._petSvc = petSvc ?? throw new ArgumentNullException(
               msg);
            InitializeComponent();         
        }

        private void btnFile_Click(object sender, EventArgs e)
        {
            string fName = "";
            string user = "";

           // ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string sqlcon = "";

            try
            {
                var build = new ConfigurationBuilder()
                   .AddJsonFile("appconfig.json", optional: false, reloadOnChange: true)
                   .AddEnvironmentVariables();
                var configuration = build.Build();

                sqlcon = configuration["ConnectionStrings:SQLConn"];

                OpenFileDialog openFileDialog1 = new OpenFileDialog
                {
                    InitialDirectory = @"C:\",
                    Title = "Browse Excel Files",

                    CheckFileExists = true,
                    CheckPathExists = true,

                    DefaultExt = "xlsx",
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    FilterIndex = 2,
                    RestoreDirectory = true,

                    ReadOnlyChecked = true,
                    ShowReadOnly = true
                };

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fName = openFileDialog1.FileName;
                }
                DataTable dt = new DataTable();


                //create a list to hold all the values
                List<string> excelData = new List<string>();

                //read the Excel file as byte array
                byte[] bin = File.ReadAllBytes(fName);
                try
                {
                    //create a new Excel package in a memorystream
                    using (MemoryStream stream = new MemoryStream(bin))
                    using (ExcelPackage excelPackage = new ExcelPackage(stream))
                    {
                        //loop all worksheets
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];
                        //check if the worksheet is completely empty
                        if (worksheet.Dimension == null)
                        {
                            // return dt;
                        }

                        //create a list to hold the column names
                        List<string> columnNames = new List<string>();

                        //needed to keep track of empty column headers
                        int currentColumn = 1;

                        //loop all columns in the sheet and add them to the datatable
                        foreach (var cell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                        {
                            string columnName = cell.Text.Trim();

                            //check if the previous header was empty and add it if it was
                            if (cell.Start.Column != currentColumn)
                            {
                                columnNames.Add("Header_" + currentColumn);
                                dt.Columns.Add("Header_" + currentColumn);
                                currentColumn++;
                            }

                            //add the column name to the list to count the duplicates
                            columnNames.Add(columnName);

                            //count the duplicate column names and make them unique to avoid the exception
                            //A column named 'Name' already belongs to this DataTable
                            int occurrences = columnNames.Count(x => x.Equals(columnName));
                            if (occurrences > 1)
                            {
                                columnName = columnName + "_" + occurrences;
                            }

                            //add the column to the datatable
                            dt.Columns.Add(columnName);

                            currentColumn++;
                        }

                        //start adding the contents of the excel file to the datatable
                        for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
                        {
                            var row = worksheet.Cells[i, 1, i, worksheet.Dimension.End.Column];
                            DataRow newRow = dt.NewRow();

                            //loop all cells in the row
                            foreach (var cell in row)
                            {
                                newRow[cell.Start.Column - 1] = cell.Text;
                            }

                            dt.Rows.Add(newRow);
                        }

                        user = Environment.UserName;

                        SqlConnection con = new SqlConnection(sqlcon);
                        con.Open();
                        SqlCommand cmd = new SqlCommand("p_InsertPets2", con);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@pets", dt));
                        cmd.Parameters.Add(new SqlParameter("@fileName", fName));
                        cmd.Parameters.Add(new SqlParameter("@userId", user));
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception tex)
                {
                   txtBox.Text = "error: " + tex.ToString();
                }
            }
            catch (Exception ex)
            {
                txtBox.Text = "error: " + ex.ToString();
            }
            
        }

        private async void Form1_Load(object sender, EventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;

          //  MessageInfo.Text = "";

            ////_userName = WindowsIdentity
            ////    .GetCurrent()
            ////    .Name
            ////    .RemoveDomain(); 

            WireUpListView();

     //       ToggleButtons(
     //           false);

            await LoadList();
        }
        private async Task<bool> LoadList()
        {
            var imported = await this._petSvc
                .FindAllFiles();

            listView1.Items.Clear();

            listView1.View = View.Details;

            // add what represents the file only, not
            // file contents
            foreach (var m in imported)
            {
                listView1.Items.Add(
                    new ListViewItem(
                        new string[] {
                            m.DataSrcID.ToString(),
                            m.fileName,
                            m.userID,
                            m.DateEntry.ToString(),
                            m.Count.ToString()}));
            }

            return true;
        }
        private void WireUpListView()
        {
            // show columns
            listView1.Columns.Add("DataSrcID");
            listView1.Columns.Add("FileName");
            listView1.Columns.Add("UserID");
            listView1.Columns.Add("DateEntry");
            listView1.Columns.Add("Count");

            listView1.GridLines = true;
            listView1.FullRowSelect = true;
            listView1.MultiSelect = false; // only 1 row at a time please

            ListViewHeaderWidth();
        }
        private void ListViewHeaderWidth()
        {
            var headerWidth = (listView1.Parent.Width - 2) / listView1.Columns.Count;

            foreach (ColumnHeader header in listView1.Columns)
                header.Width = headerWidth;
        }

        
        private void listView1_Click(object sender, EventArgs e)
        {
            var item =  listView1.SelectedItems[0].Index;
            txtBox.Text = "Your selection is: " + listView1.SelectedItems[0].SubItems[1].Text.ToString();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            int data = Int32.Parse(listView1.SelectedItems[0].SubItems[0].Text.ToString());
            getExcel(data);
        }
        private void getExcel(int data)
        {
            _petSvc.PrintExcelFile(data, FileName);
        }
    }
}
