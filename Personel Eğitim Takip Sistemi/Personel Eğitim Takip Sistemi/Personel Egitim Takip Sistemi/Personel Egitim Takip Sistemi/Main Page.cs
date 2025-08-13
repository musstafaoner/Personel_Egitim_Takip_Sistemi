using ClosedXML.Excel;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace Personel_Egitim_Takip_Sistemi
{

    public partial class Form1 : Form
    {


        // Excel dosyalarının bilgisayarda bulunduğu yerlerin tam yolu (yani adresi) burada tanımlanır
        string dosya1Yolu = @"I:\EGITIM\15 BELGE TAKİPLERİ\İMBAT MEVCUT.xlsx";
        string dosya2Yolu = @"I:\EGITIM\15 BELGE TAKİPLERİ\İMBAT MEVCUT ÖZLÜK BİLGİSİ.xlsx";
        string fotoKlasorYolu = @"I:\EGITIM\15 VERİTABANI\PERSONEL FOTOLARI";
        string dosya1 = @"I:\EGITIM\15 BELGE TAKİPLERİ\İMBAT MEVCUT.xlsx";          // Sicil No, Adı Soyadı, Departman, Görev, SGK, TC, İşe Giriş
        string dosya2 = @"I:\EGITIM\15 BELGE TAKİPLERİ\İMBAT MEVCUT ÖZLÜK BİLGİSİ.xlsx"; // Doğum tarihi, ana adı, baba adı, mezuniyet, tel no, adres, mail
        string mykKlasoru = @"I:\EGITIM\15 VERİTABANI\MYK";
        string dosyaYolu = @"I:\EGITIM\15 BELGE TAKİPLERİ\MYK KONTROL.xlsx";
        
        
        
        public Form1()
        {
            InitializeComponent(); // Arayüz bileşenlerini (buton, textbox, label vb.) başlatır


            // Butonlara tıklama olayları atanır: kullanıcı butona tıkladığında hangi işlem yapılacak?
            btnAra.Click += new EventHandler(btnAra_Click);
            btnGuncelle.Click += btnGuncelle_Click;
            btnDiplomaGoruntule.Click += btnDiplomaGoruntule_Click;
            btnKimlikGoruntule.Click += btnKimlikGoruntule_Click;

            // MYK belgeleri için butonlara olay bağlama
            btnKaynak.Click += new EventHandler(btnKaynak_Click);
            btnEndustriyel.Click += new EventHandler(btnEndustriyel_Click);
            btnHazirlik.Click += new EventHandler(btnHazirlik_Click);
            btnInsaat.Click += new EventHandler(btnInsaat_Click);
            btnVinc.Click += new EventHandler(btnVinc_Click);
            btnMakine.Click += new EventHandler(btnMakine_Click);
            btnMekanizasyon.Click += new EventHandler(btnMekanizasyon_Click);
            btnKazi.Click += new EventHandler(btnKazi_Click);
            btnPres.Click += new EventHandler(btnPres_Click);
            btnUretim.Click += new EventHandler(btnUretim_Click);

        }
        
        
        private async void btnAra_Click(object sender, EventArgs e)
        {
            txtAdSoyad.Text = "";
            txtTcKimlik.Text = "";
            txtGorevi.Text = "";
            txtDepartman.Text = "";
            txtSGKDurumu.Text = "";
            txtIseGirisTarihi.Text = "";
            txtDogumTarihi.Text = "";
            txtAnneAdi.Text = "";
            txtBabaAdi.Text = "";
            txtMezuniyet.Text = "";
            txtTelefon.Text = "";
            txtAdres.Text = "";
            txtMail.Text = "";

            string arananSicil = txtSicilNo.Text.Trim();

            if (string.IsNullOrEmpty(arananSicil))
            {
                MessageBox.Show("Lütfen sicil numarası giriniz.");
                return;
            }

            pictureBoxLoading.Visible = true; // LOADING GÖSTER

            try
            {
                // Tüm işlemleri arka planda yap (UI donmasın)
                await Task.Run(() =>
                {
                    // NOT: UI elemanlarına direkt erişemeyiz burada, Invoke kullanmamız gerekir

                    string tcNo = "";
                    bool bulunduDosya1 = false;

                    using (var wb1 = new XLWorkbook(dosya1Yolu))
                    {
                        var ws1 = wb1.Worksheet(1);
                        int lastRow1 = ws1.LastRowUsed().RowNumber();

                        for (int r = 2; r <= lastRow1; r++)
                        {
                            string sicilCell = ws1.Cell(r, 1).GetValue<string>().Trim();
                            if (sicilCell == arananSicil)
                            {
                                tcNo = ws1.Cell(r, 15).GetValue<string>().Trim();

                                // UI'ya veri yazmak için Invoke kullan
                                Invoke((MethodInvoker)(() =>
                                {
                                    txtAdSoyad.Text = ws1.Cell(r, 2).GetValue<string>();
                                    txtDepartman.Text = ws1.Cell(r, 4).GetValue<string>();
                                    txtGorevi.Text = ws1.Cell(r, 8).GetValue<string>();
                                    txtSGKDurumu.Text = ws1.Cell(r, 14).GetValue<string>();
                                    txtTcKimlik.Text = tcNo;

                                    var iseGirisCell = ws1.Cell(r, 16);
                                    if (iseGirisCell.DataType == XLDataType.DateTime)
                                        txtIseGirisTarihi.Text = iseGirisCell.GetDateTime().ToString("dd.MM.yyyy");
                                    else
                                        txtIseGirisTarihi.Text = iseGirisCell.GetValue<string>();
                                }));

                                Invoke((MethodInvoker)(() =>
                                {
                                    MykBelgeleriKontrolEt(tcNo);
                                    MykTarihleriYukle(tcNo);
                                }));

                                bulunduDosya1 = true;
                                break;
                            }
                        }
                    }

                    if (!bulunduDosya1)
                    {
                        Invoke((MethodInvoker)(() =>
                        {
                            MessageBox.Show("Personel bulunamadı (Dosya 1).");
                            Temizle();
                        }));
                        return;
                    }

                    bool bulunduDosya2 = false;
                    using (var wb2 = new XLWorkbook(dosya2Yolu))
                    {
                        var ws2 = wb2.Worksheet(1);
                        int lastRow2 = ws2.LastRowUsed().RowNumber();

                        for (int r = 2; r <= lastRow2; r++)
                        {
                            string tcCell = ws2.Cell(r, 1).GetValue<string>().Trim();
                            if (tcCell == tcNo)
                            {
                                Invoke((MethodInvoker)(() =>
                                {
                                    var dogumCell = ws2.Cell(r, 10);
                                    if (dogumCell.DataType == XLDataType.DateTime)
                                        txtDogumTarihi.Text = dogumCell.GetDateTime().ToString("dd.MM.yyyy");
                                    else
                                        txtDogumTarihi.Text = dogumCell.GetValue<string>();

                                    txtAnneAdi.Text = ws2.Cell(r, 8).GetValue<string>();
                                    txtBabaAdi.Text = ws2.Cell(r, 7).GetValue<string>();
                                    txtMezuniyet.Text = ws2.Cell(r, 5).GetValue<string>();
                                    txtTelefon.Text = ws2.Cell(r, 11).GetValue<string>();
                                    txtMail.Text = ws2.Cell(r, 12).GetValue<string>();
                                    txtAdres.Text = ws2.Cell(r, 13).GetValue<string>();
                                }));
                                bulunduDosya2 = true;
                                break;
                            }
                        }
                    }

                    // FOTO YÜKLEME
                    string[] uzantilar = { ".jpg", ".jpeg", ".png", ".bmp" };
                    bool fotoBulundu = false;

                    foreach (string uzanti in uzantilar)
                    {
                        string fotoYolu = Path.Combine(fotoKlasorYolu, tcNo + uzanti);
                        if (File.Exists(fotoYolu))
                        {
                            Invoke((MethodInvoker)(() =>
                            {
                                if (pictureBox1.Image != null)
                                {
                                    pictureBox1.Image.Dispose();
                                    pictureBox1.Image = null;
                                }
                                pictureBox1.Image = Image.FromFile(fotoYolu);
                            }));
                            fotoBulundu = true;
                            break;
                        }
                    }

                    if (!fotoBulundu)
                    {
                        Invoke((MethodInvoker)(() =>
                        {
                            if (pictureBox1.Image != null)
                            {
                                pictureBox1.Image.Dispose();
                                pictureBox1.Image = null;
                            }
                        }));
                    }
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message);
            }
            finally
            {
                pictureBoxLoading.Visible = false; // LOADING GİZLE
            }
        }

        private void btnDiplomaGoruntule_Click(object sender, EventArgs e)
        {
            string tcNo = txtTcKimlik.Text.Trim();
            if (string.IsNullOrEmpty(tcNo))
            {
                MessageBox.Show("Lütfen önce bir personel arayın.");
                return;
            }

            string dosyaYolu = $@"I:\EGITIM\15 VERİTABANI\DİPLOMA KİMLİK\DİPLOMA\{tcNo} DİP.pdf";

            if (File.Exists(dosyaYolu))
            {
                try
                {
                    var psi = new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = dosyaYolu,
                        UseShellExecute = true // PDF dosyasını varsayılan uygulamayla aç
                    };
                    System.Diagnostics.Process.Start(psi);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Dosya açılamadı: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Diploma dosyası bulunamadı.");
            }
        }



        private void btnKimlikGoruntule_Click(object sender, EventArgs e)
        {
            string tcNo = txtTcKimlik.Text.Trim();
            if (string.IsNullOrEmpty(tcNo))
            {
                MessageBox.Show("Lütfen önce bir personel arayın.");
                return;
            }

            string dosyaYolu = $@"I:\EGITIM\15 VERİTABANI\DİPLOMA KİMLİK\KİMLİK\{tcNo} KİM.pdf";

            if (File.Exists(dosyaYolu))
            {
                try
                {
                    var psi = new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = dosyaYolu,
                        UseShellExecute = true
                    };
                    System.Diagnostics.Process.Start(psi);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Kimlik dosyası açılamadı: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Kimlik dosyası bulunamadı.");
            }
        }

        private void Temizle()
        {
            txtTcKimlik.Text = "";
            txtAdSoyad.Text = "";
            txtDepartman.Text = "";
            txtGorevi.Text = "";
            txtSGKDurumu.Text = "";
            txtMezuniyet.Text = "";
            txtIseGirisTarihi.Text = "";
            txtBabaAdi.Text = "";
            txtAnneAdi.Text = "";
            txtDogumTarihi.Text = "";
            txtTelefon.Text = "";
            txtMail.Text = "";
            txtAdres.Text = "";

            if (pictureBox1.Image != null)
            {
                pictureBox1.Image.Dispose();
                pictureBox1.Image = null;
            }
        }

        private void btnGuncelle_Click(object sender, EventArgs e)
        {
            string arananSicil = txtSicilNo.Text.Trim();
            string tcNo = txtTcKimlik.Text.Trim();

            try
            {
                // === 1. DOSYA: İMBAT MEVCUT.xlsx ===
                using (var workbook1 = new XLWorkbook(dosya1))
                {
                    var ws1 = workbook1.Worksheet(1);
                    var lastRow1 = ws1.LastRowUsed().RowNumber();
                    bool bulundu1 = false;

                    for (int row = 2; row <= lastRow1; row++)
                    {
                        string cellSicil = ws1.Cell(row, 1).GetValue<string>().Trim(); // A: Sicil No
                        if (cellSicil == arananSicil)
                        {
                            ws1.Cell(row, 2).Value = txtAdSoyad.Text.Trim();        // B: Ad Soyad
                            ws1.Cell(row, 4).Value = txtDepartman.Text.Trim();      // D: Departman
                            ws1.Cell(row, 8).Value = txtGorevi.Text.Trim();         // H: Görev
                            ws1.Cell(row, 14).Value = txtSGKDurumu.Text.Trim();     // N: SGK Durumu
                            ws1.Cell(row, 15).Value = txtTcKimlik.Text.Trim();      // O: TC No
                            ws1.Cell(row, 16).Value = txtIseGirisTarihi.Text.Trim(); // P: İşe Giriş Tarihi

                            bulundu1 = true;
                            break;
                        }
                    }
                }

                // === 2. DOSYA: İMBAT MEVCUT ÖZLÜK BİLGİSİ.xlsx ===
                using (var workbook2 = new XLWorkbook(dosya2))
                {
                    var ws2 = workbook2.Worksheet(1);
                    var lastRow2 = ws2.LastRowUsed().RowNumber();
                    bool bulundu2 = false;

                    for (int row = 2; row <= lastRow2; row++)
                    {
                        string cellTc = ws2.Cell(row, 1).GetValue<string>().Trim(); // A: TC No
                        if (cellTc == tcNo)
                        {
                            ws2.Cell(row, 5).Value = txtMezuniyet.Text.Trim();     // E: Mezuniyet
                            ws2.Cell(row, 6).Value = txtIseGirisTarihi.Text.Trim(); // F: İşe Başlama (isteğe bağlı)
                            ws2.Cell(row, 7).Value = txtBabaAdi.Text.Trim();       // G: Baba Adı
                            ws2.Cell(row, 8).Value = txtAnneAdi.Text.Trim();       // H: Ana Adı
                            ws2.Cell(row, 10).Value = txtDogumTarihi.Text.Trim();  // J: Doğum Tarihi
                            ws2.Cell(row, 11).Value = txtTelefon.Text.Trim();      // K: Telefon
                            ws2.Cell(row, 12).Value = txtMail.Text.Trim();         // L: Mail
                            ws2.Cell(row, 13).Value = txtAdres.Text.Trim();        // M: Adres

                            bulundu2 = true;
                            break;
                        }
                    }
                }

                MessageBox.Show("Bilgiler başarıyla güncellendi.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message);
            }
        }

        private void MykBelgeleriKontrolEt(string tcNo)
        {
            if (string.IsNullOrWhiteSpace(tcNo))
                return;



            // Belge Adı, Klasör, Dosya adı, Buton nesnesi
            (string klasor, string dosyaAdi, Button buton)[] belgeler = new[]
            {
        ("ÇELİK KAYNAKÇISI", $"{tcNo} MYK ÇLK KAYNAK.pdf", btnKaynak),
        ("ENDÜSTRİYEL TAŞIMACI", $"{tcNo} MYK ENDÜSTRİYEL TAŞIMACI.pdf", btnEndustriyel),
        ("HAZIRLIK İŞÇİSİ", $"{tcNo} MYK HAZIRLIK İŞÇİSİ.pdf", btnHazirlik),
        ("İNŞAAT İŞÇİSİ", $"{tcNo} MYK İNŞAAT İŞÇİSİ.pdf", btnInsaat),
        ("KÖPRÜLÜ VİNÇ OPERATÖRÜ", $"{tcNo} MYK KÖPRÜLÜ VİNÇ.pdf", btnVinc),
        ("MAKİNE BAKIMCI", $"{tcNo} MYK MAKİNE BAKIMCI.pdf", btnMakine),
        ("MEKANİZASYON İŞÇİSİ", $"{tcNo} MYK MEKANİZASYON.pdf", btnMekanizasyon),
        ("MEKANİZE KAZI OPERATÖRÜ", $"{tcNo} MYK MEKANİZE KAZI.pdf", btnKazi),
        ("PRES İŞÇİSİ", $"{tcNo} MYK Pres.pdf", btnPres),
        ("ÜRETİM İŞÇİSİ", $"{tcNo} MYK ÜRETİM İŞÇİSİ.pdf", btnUretim)
    };

            foreach (var belge in belgeler)
            {
                string tamYol = Path.Combine(mykKlasoru, belge.klasor, belge.dosyaAdi);

                if (File.Exists(tamYol))
                {
                    belge.buton.Enabled = true;
                    belge.buton.Text = "Belgeyi Görüntüle";
                    belge.buton.Tag = tamYol;
                }
                else
                {
                    belge.buton.Enabled = false;
                    belge.buton.Text = "Belge Yok";
                    belge.buton.Tag = null;
                }
            }
        }

        private void btnKaynak_Click(object sender, EventArgs e)
        {
            string tcNo = txtTcKimlik.Text.Trim();
            string dosyaYolu = $@"I:\EGITIM\15 VERİTABANI\MYK\ÇELİK KAYNAKÇISI\{tcNo} MYK ÇLK KAYNAK.pdf";

            if (File.Exists(dosyaYolu))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = dosyaYolu,
                    UseShellExecute = true
                });
            }
            else
            {
                MessageBox.Show("Çelik Kaynakçısı belgesi bulunamadı.");
            }
        }

        private void btnEndustriyel_Click(object sender, EventArgs e)
        {
            string tcNo = txtTcKimlik.Text.Trim();
            string dosyaYolu = $@"I:\EGITIM\15 VERİTABANI\MYK\ENDÜSTRİYEL TAŞIMACI\{tcNo} MYK ENDÜSTRİYEL TAŞIMACI.pdf";

            if (File.Exists(dosyaYolu))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = dosyaYolu,
                    UseShellExecute = true
                });
            }
            else
            {
                MessageBox.Show("Endüstriyel Taşımacı belgesi bulunamadı.");
            }
        }

        private void btnHazirlik_Click(object sender, EventArgs e)
        {
            string tcNo = txtTcKimlik.Text.Trim();
            string dosyaYolu = $@"I:\EGITIM\15 VERİTABANI\MYK\HAZIRLIK İŞÇİSİ\{tcNo} MYK HAZIRLIK İŞÇİSİ.pdf";

            if (File.Exists(dosyaYolu))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = dosyaYolu,
                    UseShellExecute = true
                });
            }
            else
            {
                MessageBox.Show("Hazırlık İşçisi belgesi bulunamadı.");
            }
        }

        private void btnInsaat_Click(object sender, EventArgs e)
        {
            string tcNo = txtTcKimlik.Text.Trim();
            string dosyaYolu = $@"I:\EGITIM\15 VERİTABANI\MYK\İNŞAAT İŞÇİSİ\{tcNo} MYK İNŞAAT İŞÇİSİ.pdf";

            if (File.Exists(dosyaYolu))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = dosyaYolu,
                    UseShellExecute = true
                });
            }
            else
            {
                MessageBox.Show("İnşaat İşçisi belgesi bulunamadı.");
            }
        }

        private void btnVinc_Click(object sender, EventArgs e)
        {
            string tcNo = txtTcKimlik.Text.Trim();
            string dosyaYolu = $@"I:\EGITIM\15 VERİTABANI\MYK\KÖPRÜLÜ VİNÇ OPERATÖRÜ\{tcNo} MYK KÖPRÜLÜ VİNÇ.pdf";

            if (File.Exists(dosyaYolu))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = dosyaYolu,
                    UseShellExecute = true
                });
            }
            else
            {
                MessageBox.Show("Köprülü Vinç belgesi bulunamadı.");
            }
        }

        private void btnMakine_Click(object sender, EventArgs e)
        {
            string tcNo = txtTcKimlik.Text.Trim();
            string dosyaYolu = $@"I:\EGITIM\15 VERİTABANI\MYK\MAKİNE BAKIMCI\{tcNo} MYK MAKİNE BAKIMCI.pdf";

            if (File.Exists(dosyaYolu))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = dosyaYolu,
                    UseShellExecute = true
                });
            }
            else
            {
                MessageBox.Show("Makine Bakımcı belgesi bulunamadı.");
            }
        }

        private void btnMekanizasyon_Click(object sender, EventArgs e)
        {
            string tcNo = txtTcKimlik.Text.Trim();
            string dosyaYolu = $@"I:\EGITIM\15 VERİTABANI\MYK\MEKANİZASYON İŞÇİSİ\{tcNo} MYK MEKANİZASYON.pdf";

            if (File.Exists(dosyaYolu))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = dosyaYolu,
                    UseShellExecute = true
                });
            }
            else
            {
                MessageBox.Show("Mekanizasyon İşçisi belgesi bulunamadı.");
            }
        }

        private void btnKazi_Click(object sender, EventArgs e)
        {
            string tcNo = txtTcKimlik.Text.Trim();
            string dosyaYolu = $@"I:\EGITIM\15 VERİTABANI\MYK\MEKANİZE KAZI OPERATÖRÜ\{tcNo} MYK MEKANİZE KAZI.pdf";

            if (File.Exists(dosyaYolu))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = dosyaYolu,
                    UseShellExecute = true
                });
            }
            else
            {
                MessageBox.Show("Mekanize Kazı belgesi bulunamadı.");
            }
        }

        private void btnPres_Click(object sender, EventArgs e)
        {
            string tcNo = txtTcKimlik.Text.Trim();
            string dosyaYolu = $@"I:\EGITIM\15 VERİTABANI\MYK\PRES İŞÇİSİ\{tcNo} MYK Pres.pdf";

            if (File.Exists(dosyaYolu))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = dosyaYolu,
                    UseShellExecute = true
                });
            }
            else
            {
                MessageBox.Show("Pres İşçisi belgesi bulunamadı.");
            }
        }

        private void btnUretim_Click(object sender, EventArgs e)
        {
            string tcNo = txtTcKimlik.Text.Trim();
            string dosyaYolu = $@"I:\EGITIM\15 VERİTABANI\MYK\ÜRETİM İŞÇİSİ\{tcNo} MYK ÜRETİM İŞÇİSİ.pdf";

            if (File.Exists(dosyaYolu))
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                {
                    FileName = dosyaYolu,
                    UseShellExecute = true
                });
            }
            else
            {
                MessageBox.Show("Üretim İşçisi belgesi bulunamadı.");
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void MykTarihleriYukle(string tcNo)
        {

            using (var workbook = new XLWorkbook(dosyaYolu))
            {
                var ws = workbook.Worksheet("MYK KONTROL");

                int satirSayisi = ws.LastRowUsed().RowNumber();

                for (int row = 2; row <= satirSayisi; row++)
                {
                    string excelTc = ws.Cell(row, "A").GetString().Trim();

                    if (excelTc == tcNo)
                    {
                        lblHazirlikTarih.Text = GetTarih(ws.Cell(row, "G"));
                        lblMekanizasyonTarih.Text = GetTarih(ws.Cell(row, "I"));
                        lblKaynakTarih.Text = GetTarih(ws.Cell(row, "K"));
                        lblPresTarih.Text = GetTarih(ws.Cell(row, "M"));
                        lblKaziTarih.Text = GetTarih(ws.Cell(row, "O"));
                        lblVincTarih.Text = GetTarih(ws.Cell(row, "Q"));
                        lblUretimTarih.Text = GetTarih(ws.Cell(row, "S"));
                        lblEndustriyelTarih.Text = GetTarih(ws.Cell(row, "U"));
                        lblInsaatTarih.Text = GetTarih(ws.Cell(row, "W"));
                        return;
                    }
                }

                // TC bulunamazsa label'ları temizle
                lblHazirlikTarih.Text = "-";
                lblMekanizasyonTarih.Text = "-";
                lblKaynakTarih.Text = "-";
                lblPresTarih.Text = "-";
                lblKaziTarih.Text = "-";
                lblVincTarih.Text = "-";
                lblUretimTarih.Text = "-";
                lblEndustriyelTarih.Text = "-";
                lblInsaatTarih.Text = "-";
            }
        }

        private string GetTarih(IXLCell cell)
        {
            try
            {
                if (cell.IsEmpty())
                    return "-";

                // Hücre tarih formatında mı?
                if (cell.DataType == XLDataType.DateTime)
                {
                    DateTime tarih = cell.GetDateTime();
                    return tarih.ToString("dd.MM.yyyy");
                }

                // Metin içinde tarih varsa ayıkla
                string deger = cell.GetValue<string>().Trim();
                if (string.IsNullOrEmpty(deger) || deger.Contains("#YOK") || deger.Contains("####"))
                    return "-";

                // "BELGE AÇ 06.08.2029" gibi bir şeyse sondaki tarihi çek
                string[] parcalar = deger.Split(' ', '\n');
                foreach (var parca in parcalar.Reverse())
                {
                    if (DateTime.TryParse(parca, out DateTime tarih))
                        return tarih.ToString("dd.MM.yyyy");
                }

                return "-";
            }
            catch
            {
                return "-";
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
        }

        //HAYAT BOYU ÖĞRENİM BELGELERİ SAYFASINA GEÇEN BUTON
        private async void button1_Click(object sender, EventArgs e)
        {
            string girilenTcNo = txtTcKimlik.Text.Trim();

            if (!string.IsNullOrEmpty(girilenTcNo))
            {
                pictureBoxLoading2.Visible = true; // Loading göster
                pictureBoxLoading2.BringToFront(); // Öne getir

                await Task.Delay(500); // Animasyonun görünmesine fırsat ver (yarım saniye)

                Second_Page second_page = new Second_Page(girilenTcNo, this);
                second_page.Show();
                this.Hide(); // Mevcut formu gizle

                pictureBoxLoading2.Visible = false; // (istersen burada da kapatabilirsin ama form zaten gizlenmiş olacak)
            }
            else
            {
                MessageBox.Show("Lütfen TC Kimlik No giriniz.");
            }
        }

        //İŞ SAĞLIĞI VE GÜVENLİĞİ SAYFASINA GEÇEN BUTON
        private async void button2_Click(object sender, EventArgs e) 
        {
            string tc = txtTcKimlik.Text.Trim();

            if (string.IsNullOrEmpty(tc))
            {
                MessageBox.Show("Lütfen TC Kimlik No giriniz.");
                return;
            }

            pictureBoxLoading3.Visible = true; // Loading göster
            pictureBoxLoading3.BringToFront(); // Öne getir

            await Task.Delay(500);

            Thirty_Page frm3 = new Thirty_Page(tc, this);
            frm3.Show();
            pictureBoxLoading3.Visible = false;
            this.Hide();
        }
    }
}
