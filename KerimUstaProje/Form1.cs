using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.AxHost;

namespace KerimUstaProje
{
    public partial class Form1 : Form
    {
        private int rowIndex = 0;

        private static string ConnectionString =
           "Provider=MSDAORA.1;Data Source=192.168.0.150/TEST;Persist Security Info=True;Password=ifsapp;User ID=ifsapp";
        public Form1()
        {
           
            InitializeComponent();
            ComboDoldur();
            
        }


        private void button1_Click(object sender, EventArgs e)
        {
            var durum = "";
            if (comboBox1.SelectedIndex == 0)
            {
                durum = "%";
            }
            else
            {
                durum = comboBox1.SelectedItem.ToString();
            }
            using (OleDbConnection connection = new OleDbConnection(ConnectionString))
            {
                connection.Open();
                OleDbCommand cmd = new OleDbCommand("select  k.Siparis_No,    k.Yayin,    k.Redüktör_Kod,    k.Redüktör_Aciklama,   k.Adet,   k.Fiyat,    k.Toplam_Siparis_Tutar,    k.Teslim_Tarihi,   k.Durum,  k.Gecikme  ,k.onay,    k.Müsteri_Ad from YR_KERIMUSTA k where teslim_tarihi >= '" + dateTimePicker1.Value.ToShortDateString() + "' and teslim_tarihi <= '" + dateTimePicker2.Value.ToShortDateString() + "' and k.Müsteri_Ad like '" +  textBox1.Text.ToUpper() + "%" + "'and k.siparis_no like '" +  textBox2.Text.ToUpper() + "%" + "' and k.durum like '" +  durum + "%" + "'ORDER BY YAYIN", connection);
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                // DataGridView kontrolüne verileri yükle
                dataGridView1.DataSource = dt;
            }
        }

        public void ComboDoldur()
        {
            // Oracle veritabanına bağlan
            using (OleDbConnection connection = new OleDbConnection(ConnectionString))
            {
                connection.Open();

                // DISTINCT kullanarak benzersiz "STATE" değerlerini al
                OleDbCommand cmd = new OleDbCommand("select DISTINCT TRIM(LEADING ' ' FROM X.objstate) as DURUM from CUSTOMER_ORDER_LINE x ORDER BY 1", connection);
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                DataTable dt = new DataTable();

                da.Fill(dt);

                // ComboBox'a değerleri ekle
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        comboBox1.Items.Add(dt.Rows[i]["DURUM"]);
                    }
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
        }

        private void contextMenuStrip1_Click(object sender, EventArgs e)
        {
            Onayla();
        }

        private void dataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            if(e.Button == MouseButtons.Right)
            {
                dataGridView1.Rows[e.RowIndex].Selected = true;
                rowIndex = e.RowIndex;
                dataGridView1.CurrentCell = dataGridView1.Rows[rowIndex].Cells[1];
                this.contextMenuStrip1.Show(dataGridView1, e.Location);
                contextMenuStrip1.Show(Cursor.Position);
            }
        }

        public void Onayla()
        {
            string value = dataGridView1.CurrentRow.Cells[10].Value.ToString();
            MessageBox.Show(value);
            if (dataGridView1.SelectedCells.Count > 0)
            {
                int selectedRowIndex = dataGridView1.SelectedCells[0].RowIndex;
                string sip = dataGridView1.Rows[selectedRowIndex].Cells[0].Value.ToString();
                string yay = dataGridView1.Rows[selectedRowIndex].Cells[1].Value.ToString();
                string reduk = dataGridView1.Rows[selectedRowIndex].Cells[2].Value.ToString();
                // Veritabanı bağlantısı
                if (value == "Onaysız")
                { 
                    using (OleDbConnection connection = new OleDbConnection(ConnectionString))
                    {
                        connection.Open();
                        string selectQuery = $"SELECT ONAY FROM Yr_Ustakerim_Proje2_Tab WHERE SIPARIS_NO = '{sip}' AND YAYIN = '{yay}' AND REDUKTOR_KOD = '{reduk}'";
                        using (OleDbCommand selectCommand = new OleDbCommand(selectQuery, connection))
                        {
                            OleDbDataReader reader = selectCommand.ExecuteReader();
                            if (reader.Read())
                            {
                                // Eğer kayıt varsa ONAY durumunu güncelle
                                string updateQuery = $"UPDATE Yr_Ustakerim_Proje2_Tab SET ONAY = 'Onaylı',ONAY_TARIH = ? WHERE SIPARIS_NO = '{sip}' AND YAYIN = '{yay}' AND REDUKTOR_KOD = '{reduk}'";
                                using (OleDbCommand updateCommand = new OleDbCommand(updateQuery, connection))
                                {
                                    
                                    updateCommand.Parameters.AddWithValue("ONAY_TARIH", DateTime.Now);
                                    int rowsAffected = updateCommand.ExecuteNonQuery();

                                    if (rowsAffected > 0)
                                    {
                                        dataGridView1.Rows[selectedRowIndex].Cells["ONAY"].Value = "Onaylı";
                                        MessageBox.Show("Onay durumu güncellendi.");
                                    }
                                    else
                                    {
                                        MessageBox.Show("Onay durumu güncellenirken hata oluştu.");
                                    }
                                }
                           
                            }
                            else
                            {
                                // Eğer kayıt yoksa yeni kayıt ekle
                                string insertQuery = $"INSERT INTO Yr_Ustakerim_Proje2_Tab (SIPARIS_NO, YAYIN, REDUKTOR_KOD, ONAY,ONAY_TARIH) VALUES ('{sip}', '{yay}', '{reduk}', 'Onaylı',?)";
                                using (OleDbCommand insertCommand = new OleDbCommand(insertQuery, connection))
                                {
                                   
                                    dataGridView1.Rows[selectedRowIndex].Cells["ONAY"].Value = "Onaylı";
                                    insertCommand.Parameters.AddWithValue("ONAY_TARIH", DateTime.Now);
                                    int rowsAffected = insertCommand.ExecuteNonQuery();

                                    if (rowsAffected > 0)
                                    {
                                        dataGridView1.Rows[selectedRowIndex].Cells["ONAY"].Value = "Onaylı";
                                        MessageBox.Show("Yeni Onay eklendi.");
                                    }
                                    else
                                    {
                                        MessageBox.Show("Yeni Onay eklenirken hata oluştu.");
                                    }
                                }
                             
                            }
                        }
                    }
                }
                else if (value == "Onaylı")
                {
                    if (dataGridView1.SelectedCells.Count > 0)
                    {
                       // int selectedRowIndex = dataGridView1.SelectedCells[0].RowIndex;
                        string orderNo = dataGridView1.Rows[selectedRowIndex].Cells[0].Value.ToString();
                        string lineNo = dataGridView1.Rows[selectedRowIndex].Cells[1].Value.ToString();
                        string catalogNo = dataGridView1.Rows[selectedRowIndex].Cells[2].Value.ToString();
                        // Veritabanı bağlantısı
                        
                        using (OleDbConnection connection = new OleDbConnection(ConnectionString))
                        {
                            connection.Open();
                            string updateQuery = $"UPDATE Yr_Ustakerim_Proje2_Tab SET ONAY = 'Onaysız',ONAY_TARIH = ? WHERE SIPARIS_NO = '{sip}' AND YAYIN = '{yay}' AND REDUKTOR_KOD = '{reduk}'";
                            using (OleDbCommand updateCommand = new OleDbCommand(updateQuery, connection))
                            {

                                updateCommand.Parameters.AddWithValue("ONAY_TARIH", DateTime.Now);
                                int rowsAffected = updateCommand.ExecuteNonQuery();

                                if (rowsAffected > 0)
                                {
                                    dataGridView1.Rows[selectedRowIndex].Cells["ONAY"].Value = "Onaysız";
                                    MessageBox.Show("Onay durumu güncellendi.");
                                }
                                else
                                {
                                    MessageBox.Show("Onay durumu güncellenirken hata oluştu.");
                                }
                            }
                          
                        }
                    }
                    else
                    {
                        MessageBox.Show("Lütfen bir satır seçin.");
                    }
                }
            }
            else
            {
                MessageBox.Show("Lütfen bir satır seçin.");
            }
        }
        
        //excele çıktı alma metodu
        public static void Excel_Disa_Aktar(DataGridView dataGridView1)
        {
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            app.Visible = true;
            worksheet = workbook.Sheets["Sayfa1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "Excel Dışa Aktarım";
            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }
            Microsoft.Office.Interop.Excel.Range usedRange = worksheet.UsedRange;
            usedRange.Columns.AutoFit();
            usedRange.Rows.AutoFit();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Excel_Disa_Aktar(dataGridView1);
        }


    }
}