using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
            Crypto crypto = new Crypto();
            var name = crypto.DecryptCode("UDsM886D8EjbjmI19DyO7g==");
            MessageBox.Show(name);
        }
    }
}
