using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat;

namespace ExcelEsitleme
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string DosyaYolu = "";
        string DosyaAdi = "";
        string DosyaYolu2 = "";
        string DosyaAdi2 = "";
        string SayfaAdi1 = "";
        string SayfaAdi2 = "";
        string araAlanAdi = "";
        string exePath = "";
        string excelPath = "";

        private void sayfaAdiAta()
        {
            SayfaAdi1 = textBox3.Text;
            SayfaAdi2 = textBox4.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            //file.Filter = "Excel Dosyası |*.xlsx|*.xls";
            file.Filter = "Excel Dosyası (*.xlsx,*.xls) | *.xlsx; *.xls";

            file.RestoreDirectory = true;
            file.CheckFileExists = false;
            file.Title = "Excel Dosyası Seçiniz..";

            if (file.ShowDialog() == DialogResult.OK)
            {
                DosyaYolu = file.FileName;
                DosyaAdi = file.SafeFileName;
                textBox1.Text = DosyaYolu;

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            //file.Filter = "Excel Dosyası |*.xlsx|*.xls";
            file.Filter = "Excel Dosyası (*.xlsx,*.xls) | *.xlsx; *.xls";

            file.RestoreDirectory = true;
            file.CheckFileExists = false;
            file.Title = "Excel Dosyası Seçiniz..";

            if (file.ShowDialog() == DialogResult.OK)
            {
                DosyaYolu2 = file.FileName;
                DosyaAdi2 = file.SafeFileName;
                textBox2.Text = DosyaYolu2;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            excelPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            exePath = System.Windows.Forms.Application.StartupPath;
            sayfaAdiAta();
            araAlanAdi = textBox5.Text;
            string ConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + exePath + "\\bilgi1.accdb";
            OleDbConnection conn = new OleDbConnection(ConnString);
            conn.Open();
            string sql = @"select * into ana from [Excel 8.0;HDR=YES;DATABASE=" + DosyaYolu + "].[" + SayfaAdi1 + "$] s;";
            string sql2 = @"select * into yavru from [Excel 8.0;HDR=YES;DATABASE=" + DosyaYolu2 + "].[" + SayfaAdi2 + "$] s;";

            OleDbCommand cmd = new OleDbCommand();
            cmd.Connection = conn;
            cmd.CommandText = sql;
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql2;
            cmd.ExecuteNonQuery();

            string no = araAlanAdi;
            string sql3 = "SELECT * FROM [" + exePath + "\\bilgi1.accdb].ana a INNER JOIN [" + exePath + "\\bilgi1.accdb].yavru b ON a." + no + "  = b." + no;
            cmd.CommandText = sql3;


            OleDbDataAdapter da = new OleDbDataAdapter(sql3, ConnString);

            try
            {

                DataSet ds = new DataSet("Test");
                System.Windows.Forms.Application.DoEvents();


                MessageBox.Show("Excel Dosyasına Yüklendi.");

                da.Fill(ds, "Test");
                dtGrid.DataSource = ds.Tables["Test"].DefaultView;
                dtGrid.AutoGenerateColumns = true;

                da.AcceptChangesDuringFill = true;
                da.Dispose();




                ExportDataSetToExcel(ds);

                string dropTableAna = "DROP TABLE [ana]";
                string dropTableYavru = "DROP TABLE [yavru]";

                OleDbCommand cmdDropAna = new OleDbCommand(dropTableAna, conn);
                cmdDropAna.ExecuteNonQuery();

                OleDbCommand cmdDropYavru = new OleDbCommand(dropTableYavru, conn);
                cmdDropYavru.ExecuteNonQuery();




            }
            catch (Exception)
            {
                string dropTableAna = "DROP TABLE [ana]";
                string dropTableYavru = "DROP TABLE [yavru]";

                OleDbCommand cmdDropAna = new OleDbCommand(dropTableAna, conn);
                cmdDropAna.ExecuteNonQuery();

                OleDbCommand cmdDropYavru = new OleDbCommand(dropTableYavru, conn);
                cmdDropYavru.ExecuteNonQuery();
                MessageBox.Show("Kolon isimlerinde hata var. Kolon isimlerinde boşluk olamaz ve Türkçe karakterler bulunamaz.");
            }
            finally
            {
                conn.Close();
            }


        }

        private void ExportDataSetToExcel(DataSet ds)
        {
            string excelFileName = excelPath + "\\BilgiGetir.xls";
            ExcelLibrary.DataSetHelper.CreateWorkbook(excelFileName, ds);

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //
        }

    }
}
