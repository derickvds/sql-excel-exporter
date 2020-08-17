using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace SqlExcelExporter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void ExportBtn_Click(object sender, EventArgs e)
        {

            _Application excel = new _Excel.Application();
            string cs = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            StringBuilder sb = new StringBuilder();
            Crypto crypto = new Crypto();
            Workbook wb;
            Worksheet ws;

            wb = excel.Workbooks.Add();
            ws = excel.Sheets["Sheet1"];
            int row = 1;

            using (SqlConnection con = new SqlConnection(cs))
            {
                SqlDataAdapter da = new SqlDataAdapter("SELECT [PatientNumber],[FirstName],[LastName],[ContactNumber],[Address] FROM [Patients]", con);
                DataSet ds = new DataSet();
                da.Fill(ds);

                ds.Tables[0].TableName = "Patients";

                foreach (DataRow patientRow in ds.Tables["Patients"].Rows)
                {
                    int patientNumber = Convert.ToInt32(patientRow["PatientNumber"]);
                    ws.Cells[row, 1].Value2 = patientNumber.ToString();
                    ws.Cells[row, 2].Value2 = crypto.DecryptCode(patientRow["FirstName"].ToString());
                    ws.Cells[row, 3].Value2 = crypto.DecryptCode(patientRow["LastName"].ToString());
                    ws.Cells[row, 4].Value2 = crypto.DecryptCode(patientRow["ContactNumber"].ToString());
                    ws.Cells[row, 5].Value2 = crypto.DecryptCode(patientRow["Address"].ToString());

                    row++;
                }
            }

            
            wb.SaveAs(@"C:\ExportedData\PatientsList.xls");
            wb.Close();
            MessageBox.Show("Done");
        }
    }
}
