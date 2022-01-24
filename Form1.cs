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
using System.IO;
using ExcelDataReader;
using Z.Dapper.Plus;

namespace SQLImportExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            DataTable dt = tableCollection[comboBox1.SelectedItem.ToString()];
            //dataGridView2.DataSource = dt;
            if(dt != null)
            {
                List<siparis> siparisler = new List<siparis>();
                for(int i = 0; i < dt.Rows.Count; i++)
                {
                    siparis sipariss = new siparis();
                    sipariss.ord_code = dt.Rows[i]["ord_code"].ToString();
                    sipariss.SSCC = dt.Rows[i]["SSCC"].ToString();
                    siparisler.Add(sipariss);
                }
                siparisBindingSource.DataSource = siparisler;
            }
        }

        DataTableCollection tableCollection;
        //SqlConnection baglanti = new SqlConnection("Data Source = 192.168.167.195; Initial Catalog = LV3; User ID = sa; Password = Password1");
        private void button2_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog ope = new OpenFileDialog();
            ope.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (ope.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = ope.FileName;
                using (var stream = File.Open(ope.FileName, FileMode.Open, FileAccess.Read))
                {
                    using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                        });
                        tableCollection = result.Tables;
                        comboBox1.Items.Clear();
                        foreach (DataTable table in tableCollection)
                            comboBox1.Items.Add(table.TableName);
                    }
                }
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            try
            {
                string connectionstring ="Server = 192.168.167.195;Database=LV3;User ID=sa;Password=Password1;";
                DapperPlusManager.Entity<siparis>().Table("HRZ_CompanyOrderListEcom");
                List<siparis> siparisler = siparisBindingSource.DataSource as List<siparis>;
                if (siparisler != null)
                {
                    using (IDbConnection db = new SqlConnection(connectionstring))
                    {
                        db.BulkInsert(siparisler);
                    }
                }
                MessageBox.Show("Finish");
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
                //baglanti.Open();
                //SqlCommand kayitekle = new SqlCommand("insert into HRZ_CompanyOrderListEcom_cy (ord_code,SSCC) values (@p1,@p2)", baglanti);
                //kayitekle.Parameters.AddWithValue("@p1", textBox1.Text);
                //kayitekle.Parameters.AddWithValue("@p2", textBox2.Text);

                //kayitekle.ExecuteNonQuery();
                //MessageBox.Show("Başarılı");
                //baglanti.Close();
                //}
