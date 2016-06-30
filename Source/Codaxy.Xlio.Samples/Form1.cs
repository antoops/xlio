using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace Codaxy.Xlio.Samples
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var workbook = new Workbook();
            var sheet = workbook.Sheets.AddSheet("Demo");
            sheet["A1"].Value = "Hello World";
            sheet["A1"].Style.Font.Size = 20;
            workbook.Save(@"FirstWorkbook.xlsx");


 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var workbook = Workbook.Load("FirstWorkbook.xlsx");
            var sheet = workbook.Sheets[0];

            
            int i = sheet.Data.Count; //getRowCount();
            i++;

            sheet["A"+i.ToString()].Value = "Hello World";
            sheet["A"+i.ToString()].Style.Font.Size = 20;
            workbook.Save(@"FirstWorkbook.xlsx");

        }

        private int getRowCount()
        {
            var fileName = string.Format("{0}\\FirstWorkbook.xlsx", Directory.GetCurrentDirectory());
            var connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0\";";
            var adapter = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", connectionString);
            var ds = new DataSet();
            adapter.Fill(ds, "Sheet1");
            // var connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fileName);
            //  var connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 12.0;HDR=YES;IMEX=1",fileName);

            // var connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";";

            //var connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0\";";


            var adapter = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", connectionString);
            var ds = new DataSet();

            adapter.Fill(ds,"Sheet1");

            DataTable data = ds.Tables["Sheet1"];

            return data.Rows.Count;
            //MessageBox.Show(data.Rows[778]["Date"].ToString());
            //// create date time 2015-09-02 00:00:00.000
            //DateTime dt = new DateTime(2016, 6, 28, 00, 00, 00, 000);
            //// MessageBox.Show(String.Format("{0:dd-MM-yyyy hh:mm:ss tt}", dt));
            //string currentDate = String.Format("{0:M/d/yyyy}", dt);

            //DataView dv = new DataView(data);
            //dv.Sort = "Date";
            //DataRowView[] results = dv.FindRows(currentDate);


            //if (results.Length <= 0)
            //    MessageBox.Show("no records");
            //else
            //{
            //    foreach (DataRowView dr in results)
            //        ddlTasks.Items.Add(dr["Task"]);
            //}


        }
    }
}
