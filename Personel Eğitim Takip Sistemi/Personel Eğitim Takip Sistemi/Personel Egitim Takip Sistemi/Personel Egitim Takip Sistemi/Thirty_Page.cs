using ClosedXML.Excel;
using iTextSharp.text.pdf;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Personel_Egitim_Takip_Sistemi
{
    public partial class Thirty_Page : Form
    {
        private Form1 _mainPage;
        private string tcNo;
        string dosyaYolu = @"I:\EGITIM\15 BELGE TAKİPLERİ\GÖREV TANIM FORMLARI KONTROL .xlsx";

        public Thirty_Page(string gelenTc, Form1 mainPage)
        {
            InitializeComponent();
            this._mainPage = mainPage;
            tcNo = gelenTc.Trim(); // TC'yi al ve boşlukları temizle


            // 2024'ten 2040'a kadar yıllar döngüyle dönülür
            for (int yil = 2024; yil <= 2040; yil++)
            {
                string labelAd1 = $"lblisg{yil % 100}1";
                string labelAd2 = $"lblisg{yil % 100}2";

                Label label1 = this.Controls.Find(labelAd1, true).FirstOrDefault() as Label;
                Label label2 = this.Controls.Find(labelAd2, true).FirstOrDefault() as Label;

                if (label1 != null && label2 != null)
                {
                    ExcelVerisiYukle(yil.ToString(), label1, label2);
                }
            }

            //  Burada veri çekimi yapılır.
            VerileriGetir(tcNo);
            //İŞ EKİPMANLARI EĞİTİMİNİ AÇARKEN BURASI DA AÇILACAK
            //IsEkipmanıEgitimleriniYukle(tcNo);
        }

        public void VerileriGetir(string gelenTc)
        {
            try
            {
                using (var workbook = new XLWorkbook(dosyaYolu))
                {
                    var ws = workbook.Worksheet(1); // İstenilen sayfa adıysa ad ile kullanılabilir
                    int sonSatir = ws.LastRowUsed().RowNumber();

                    for (int i = 2; i <= sonSatir; i++) // Başlık satırını atla
                    {
                        string excelTc = ws.Cell(i, 1).GetString().Trim();

                        if (string.Equals(excelTc, gelenTc.Trim(), StringComparison.OrdinalIgnoreCase))
                        {
                            textBox1.Text = ws.Cell(i, 7).GetString();
                            textBox2.Text = ws.Cell(i, 8).GetString();
                            textBox3.Text = ws.Cell(i, 9).GetString();
                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message);
            }
        }
        public void ExcelVerisiYukle(string yil, Label hedefLabel1, Label hedefLabel2)
        {
            try
            {
                string klasorYolu = $@"I:\EGITIM\14 İSG HATIRLATMA EĞİTİMLERİ\{yil} İSG HATIRLATMA EĞİTİMİ";
                string dosyaAdi = $"{yil} HATIRLATMA.xlsx";
                string tamYol = Path.Combine(klasorYolu, dosyaAdi);

                if (!File.Exists(tamYol))
                {
                    hedefLabel1.Text = "-";
                    hedefLabel2.Text = "-";
                    return;
                }

                using (var wb = new XLWorkbook(tamYol))
                {
                    var ws = wb.Worksheet(2);
                    var satirlar = ws.RangeUsed().RowsUsed().Skip(1);

                    foreach (var satir in satirlar)
                    {
                        string excelTc = satir.Cell(1).GetValue<string>().Trim();

                        if (excelTc == tcNo)
                        {
                            string cell9 = satir.Cell(9).GetValue<string>().Trim();
                            string cell10 = satir.Cell(10).GetValue<string>().Trim();

                            string egitim1gun = string.IsNullOrWhiteSpace(cell9) ? "-" : cell9.Split(' ')[0];
                            string egitim2gun = string.IsNullOrWhiteSpace(cell10) ? "-" : cell10.Split(' ')[0];

                            hedefLabel1.Text = egitim1gun;
                            hedefLabel2.Text = egitim2gun;
                            return;
                        }
                    }

                    hedefLabel1.Text = "TC bulunamadı";
                    hedefLabel2.Text = "TC bulunamadı";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu:\n" + ex.Message);
                hedefLabel1.Text = "HATA";
                hedefLabel2.Text = "HATA";
            }
        }

        //İŞ EKİPMANI EĞİİTMLERİNİ YÜKLERKEN AÇILACAK
        //private void IsEkipmanıEgitimleriniYukle(string gelenTc)
        //{
        //    try
        //    {
        //        string klasorYolu = $@"I:\EGITIM\3 SERTİFİKALAR + OFK+K.DÖNÜŞ\8- İŞ EKİPMALARI EĞİTİMİ SERTİFİKALAR";
        //        string dosyaAdi = $"İŞ EKİPMANI EĞİTİMİ ALANLAR.xlsx";
        //        string tamYol = Path.Combine(klasorYolu, dosyaAdi);


        //        using (var wb = new XLWorkbook(tamYol))
        //        {
        //            var ws1 = wb.Worksheet(2);
        //            var satirlar1 = ws1.LastRowUsed().RowNumber();

        //            for (int i = 2; i <= satirlar1; i++)
        //            {
        //                string excelTc = ws1.Cell(i, 1).GetString().Trim();

        //                if (excelTc == tcNo)
        //                {

        //                    txtGorevAdı.Text = ws1.Cell(i, 4).GetString();
        //                    txtEgitimTarihi.Text = ws1.Cell(i, 9).GetDateTime().ToString("dd.MM.yyyy");
        //                    return;

        //                }
        //            }
        //            MessageBox.Show("TC bulunamadı.");

        //            txtGorevAdı.Text = "Görev Bulunamadı.";
        //            txtEgitimTarihi.Text = "Tarih Bulunamadı";
        //        }
        //    }

        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("Hata oluştu:\n" + ex.Message);
        //        txtGorevAdı.Text = "HATA";
        //        txtEgitimTarihi.Text = "HATA";
        //    }
        //}

        private void button1_Click(object sender, EventArgs e)
        {
            _mainPage.Show();
            this.Hide();
        }

        private void Thirty_Page_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

    }
}
