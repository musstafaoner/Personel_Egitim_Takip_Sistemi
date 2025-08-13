using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Personel_Egitim_Takip_Sistemi
{
    public partial class Second_Page : Form
    {
        
        private string tcNo;
        private Form1 _mainPage;
        public Second_Page(string gelenTcNo, Form1 mainPage)
        {
            InitializeComponent();
            this.tcNo = gelenTcNo;
            this._mainPage = mainPage;
            HayatBoyuBelgeleriKontrolEt(tcNo);


            // HAYAT BOYU ÖĞRENME BUTON EVENTLERİ
            this.btnCevher.Click += new EventHandler(this.btnCevher_Click);
            this.btnDemiryolu.Click += new EventHandler(this.btnDemiryolu_Click);
            this.btnElle.Click += new EventHandler(this.btnElle_Click);
            this.btnForklift.Click += new EventHandler(this.btnForklift_Click);
            this.btnGaleri.Click += new EventHandler(this.btnGaleri_Click);
            this.btnIsMak.Click += new EventHandler(this.btnIsMak_Click);
            this.btnHijyen.Click += new EventHandler(this.btnHijyen_Click);
            this.btnKaynakci.Click += new EventHandler(this.btnKaynakci_Click);
            this.btnKazici.Click += new EventHandler(this.btnKazici_Click);
            this.btnKlasik.Click += new EventHandler(this.btnKlasik_Click);
            this.btnBant.Click += new EventHandler(this.btnBant_Click);
            this.btnZincirli.Click += new EventHandler(this.btnZincirli_Click);
            this.btnManevraci.Click += new EventHandler(this.btnManevraci_Click);
            this.btnMekanik.Click += new EventHandler(this.btnMekanik_Click);
            this.btnNezaretci.Click += new EventHandler(this.btnNezaretci_Click);
            this.btnMetal.Click += new EventHandler(this.btnMetal_Click);
            this.btnMonoray.Click += new EventHandler(this.btnMonoray_Click);
            this.btnNakliyatBak.Click += new EventHandler(this.btnNakliyatBak_Click);
            this.btnNakliyatU.Click += new EventHandler(this.btnNakliyatU_Click);
            this.btnNezaretci2.Click += new EventHandler(this.btnNezaretci2_Click);
            this.btnPresci.Click += new EventHandler(this.btnPresci_Click);
            this.btnRamble.Click += new EventHandler(this.btnRamble_Click);
            this.btnSondaj.Click += new EventHandler(this.btnSondaj_Click);
            this.btnTamburB.Click += new EventHandler(this.btnTamburB_Click);
            this.btnTamburO.Click += new EventHandler(this.btnTamburO_Click);
            this.btnTulumba.Click += new EventHandler(this.btnTulumba_Click);
            this.btnYuruyenH.Click += new EventHandler(this.btnYuruyenH_Click);
            this.btnYuruyenS.Click += new EventHandler(this.btnYuruyenS_Click);
        }

        private void BelgeAc(string dosyaYolu)
        {
            if (File.Exists(dosyaYolu))
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = dosyaYolu,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Belge açılamadı: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Belge bulunamadı.");
            }
        }

        public void HayatBoyuBelgeleriKontrolEt(string tcNo)
        {
            if (string.IsNullOrWhiteSpace(tcNo))
                return;

            string klasorBase = @"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM";

            (string altKlasor, Button buton)[] belgeler = new[]
            {
        ("CEVHER", btnCevher),
        ("DEMİRYOLU", btnDemiryolu),
        ("ELLE TAŞIMA", btnElle),
        ("FORKLİFT", btnForklift),
        (@"GAM VE İŞ MAK.OP\GALERİ AÇMA", btnGaleri),
        (@"GAM VE İŞ MAK.OP\İŞ MAK", btnIsMak),
        ("HİJYEN MUTFAK", btnHijyen),
        ("KAYNAKÇI", btnKaynakci),
        ("KAZI", btnKazici),
        ("KLASİK SİSTEM", btnKlasik),
        (@"KONVEYÖR\BANT KONVEYÖR", btnBant),
        (@"KONVEYÖR\ZİNCİRLİ KONVEYÖR", btnZincirli),
        ("MANEVRACI", btnManevraci),
        ("MEK.TAM.B", btnMekanik),
        ("MEKANİK NEZARETÇİ", btnNezaretci),
        ("METAL BYM", btnMetal),
        ("MONORAY", btnMonoray),
        ("NAKLİYAT BAK", btnNakliyatBak),
        ("NAKLİYAT Ü.ÇAL", btnNakliyatU),
        ("NEZARETÇİ", btnNezaretci2),
        ("PRES", btnPresci),
        ("RAMBLE", btnRamble),
        ("SONDAJ", btnSondaj),
        ("TAMBURLU KESİCİ YÜKLEYİCİ BAKIMCISI", btnTamburB),
        ("TAMBURLU KESİCİ YÜKLEYİCİ OPERATÖRÜ", btnTamburO),
        ("TULUMBA", btnTulumba),
        ("YÜRÜYEN TAHKİMAT HİDROLİK BAKIMCISI", btnYuruyenH),
        ("YÜRÜYEN TAHKİMAT SÜRÜCÜSÜ", btnYuruyenS)
    };

            foreach (var belge in belgeler)
            {
                string dosyaYolu = Path.Combine(klasorBase, belge.altKlasor, $"{tcNo}.pdf");

                if (File.Exists(dosyaYolu))
                {
                    belge.buton.Enabled = true;
                    belge.buton.Text = "Belgeyi Görüntüle";
                    belge.buton.Tag = dosyaYolu;
                }
                else
                {
                    belge.buton.Enabled = false;
                    belge.buton.Text = "Belge Yok";
                    belge.buton.Tag = null;
                }
            }
        }

        private void btnCevher_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\CEVHER\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnDemiryolu_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\DEMİRYOLU\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnElle_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\ELLE TAŞIMA\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnForklift_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\FORKLİFT\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnGaleri_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\GAM VE İŞ MAK.OP\GALERİ AÇMA\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnIsMak_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\GAM VE İŞ MAK.OP\İŞ MAK\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnHijyen_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\HİJYEN MUTFAK\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnKaynakci_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\KAYNAKÇI\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnKazici_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\KAZI\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnKlasik_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\KLASİK SİSTEM\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnBant_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\KONVEYÖR\BANT KONVEYÖR\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnZincirli_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\KONVEYÖR\ZİNCİRLİ KONVEYÖR\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnManevraci_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\MANEVRACI\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnMekanik_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\MEK.TAM.B\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnNezaretci_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\MEKANİK NEZARETÇİ\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnMetal_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\METAL BYM\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnMonoray_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\MONORAY\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnNakliyatBak_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\NAKLİYAT BAK\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnNakliyatU_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\NAKLİYAT Ü.ÇAL\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnNezaretci2_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\NEZARETÇİ\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnRamble_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\RAMBLE\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnSondaj_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\SONDAJ\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnTamburB_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\TAMBURLU KESİCİ YÜKLEYİCİ BAKIMCISI\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnTamburO_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\TAMBURLU KESİCİ YÜKLEYİCİ OPERATÖRÜ\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnTulumba_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\TULUMBA\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnYuruyenH_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\YÜRÜYEN TAHKİMAT HİDROLİK BAKIMCISI\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void btnYuruyenS_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\YÜRÜYEN TAHKİMAT SÜRÜCÜSÜ\{tcNo}.pdf";
            BelgeAc(path);
        }
        private void btnPresci_Click(object sender, EventArgs e)
        {
            string path = $@"I:\EGITIM\15 VERİTABANI\HAYAT BOYU ÖĞRENİM\PRES\{tcNo}.pdf";
            BelgeAc(path);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            _mainPage.Show(); // eski sayfayı geri getir
            this.Hide();
        }

        private void Second_Page_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
    }

}


