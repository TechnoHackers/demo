using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace CHECKCIE
{
    public partial class Form1 : Form
    {
        string ruta;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void bunifuImageButton1_Click(object sender, EventArgs e)
        {
            string conexion = "Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source= "+ruta+" ; Extended Properties =\"Excel 8.0;HDR=Yes\"";
            OleDbDataReader read;
            OleDbConnection con = new OleDbConnection(conexion);
            con.Open();
            OleDbCommand cmd = new OleDbCommand("select * from [Hoja1$]", con);
            OleDbDataAdapter dta = new OleDbDataAdapter(cmd);
            DataTable dt = new DataTable();
            dta.Fill(dt);
            dataGridView1.DataSource = dt;
            read = cmd.ExecuteReader();
            while (read.Read())
            {
                listBox1.Items.Add(read["sexo"]);
            }
            
        }

        private void bunifuFlatButton1_Click(object sender, EventArgs e)
        {
            OpenFileDialog archivo = new OpenFileDialog();
            if (archivo.ShowDialog() == DialogResult.OK)
            {
                ruta = archivo.FileName;
            }
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            Microsoft.Office.Interop.Excel.Workbooks books = excelApp.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook sheet = books.Open(ruta);
        }
    }
}
