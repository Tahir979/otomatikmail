using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;
using System.Text.RegularExpressions;
using System.IO;
//using Excel = Microsoft.Office.Interop.Excel; //microsoft office.14.object library, sonra nugetten microsoft office interop'u ekleyip yapacağız
//bu aynı isimlere ikinci kez mail göndermemeyi ve ayıklatmayı bu programa da aktarayım ya

//mevcut mail sistemde kalmaya devam ediyor doğrulama kodu hüpt gmail hesabına gönderilmedi (mükemmel olacak bu varya ya; o mail adresi ve şifre değişebilir onu ayarlamak lazım)(ama o kısımda bu kutucukları boş doldurmalarına asla izin verme yoksa kötü olur valla)
//hatalı mail illaki olacaktır onlara nasıl bir çare bulmak lazım hadi atarken gitti insanlar bilemez ki o durumda ne yapacaklarını
//klasör gönderebileck mi acaba; her dosya türünden yollayabilecek mi
//splashscreen yapsam mı ki lan acaba gsgsf
//yok abi bulamadım valla o progress bar marquee değeri neden öyle oluyor, sanırsam splash screen multithread filan ögrenmem lazım


//diyelim dosyayı yükledi ve sonra dosyanın ismini değiştirdi, o zaman program hata verecektir yolu bulamadığı için dosyanın dikkat etmek lazım bu tarz programın hata verebileceği ince durumlara(ufak tefek ince detaylar kısmına ekleyelim bunu)

namespace Otomatik_Mail_Atma //Mail İmzası(bunu nasıl ekletebiliriz ki acaba ya bi düşünmek lazım üstüne (acaba otomatik olarak en alta ekleyebilir miyim ki onu bir şekilde
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        //Void kodlarının açıklaması aşağıdaki "region" yapısının içinde anlatılmıştır, kafanıza takılan yer olursa oradan faydalanabilirsiniz
        public Form1()
        {
            InitializeComponent();
        } //
        public static string TersCevir(string metin)
        {
            string sonuc = "";
            for (int i = metin.Length - 1; i >= 0; i--)
                sonuc += metin[i];
            return sonuc;
        } //A
        void Emailgonder()
        {
            if(metroLabel11.Text == "0") //eğer tek başına if olarak kalsaydı daha sonra alttan try catch komutuna yine devam edecektir ama şimdi else koyduğun için eğer veri yoksa durucak ve try içindeki işlemleri yapmaya girişmeyecek. eğer if şartı sağlamasaydı bu sefer direkt try yapısına inmiş olacaktı ama istemediğimiz şekilde yani veri girişi yapılmış mı yapılmamış mı bunun kontrolünü yapmadan ilerleyecekti
            {
                MessageBox.Show("Herhangi bir veri girişi bulunamadı","Hata",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            else
            {
                SmtpClient sc = new SmtpClient
                {
                    Port = 587,
                    Timeout = 3600000, //tamam la çalışıyormuş bu gerçekten
                    Host = "smtp.gmail.com",
                    EnableSsl = true,
                    Credentials = new NetworkCredential(Settings.Default.mail, Settings.Default.sifre)
                };
                MailMessage mail = new MailMessage
                {
                    From = new MailAddress(Settings.Default.mail, "Hacettepe Üniversitesi Psikoloji Topluluğu (HÜPT)")
                };
                mail.To.Add(txt_email.Text.ToString());
                mail.Subject = txt_baslik.Text; mail.IsBodyHtml = true; mail.Body = txt_hitap.Text + " " + txt_isim.Text.ToString() + "," + "<br/> <br/>" + txt_icerik2.Text;
                mail.Attachments.Clear();
                if (checkedListBox1.Items.Count != 0)
                {
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        if (checkedListBox1.GetItemCheckState(i) == CheckState.Checked)
                        {
                            mail.Attachments.Add(new Attachment(checkedListBox2.Items[i].ToString()));
                        }
                    }
                }

                sc.Send(mail);
            }
        } //B
        void Mbcheck()
        {
            if (listBox5.Items.Count == 0)
            {
                metroLabel6.Text = "0.00 MB";
            }
            else
            {
                double bayt = 0;
                for (int i = 0; i < listBox5.Items.Count; i++)
                {
                    bayt += Convert.ToInt32(listBox5.Items[i].ToString());
                }
                double x = 0.00000095367432;
                double mb = bayt * x;
                mb = Math.Round(mb, 2);

                metroLabel6.Text = mb + " MB";
            }
        } //C
        void Check1()
        {
            string metin;
            int x = 0;

            for (int i = 0; i < checkedListBox1.Items.Count; i++) //hmm bak gördün mü; burayı sabit değer yerine değişken değerler alınca bak nasıl da düzeldi int g = checkedListBox1.Items.Count olabilir ama o sadece baştaki sayısı neyse onu kabul eder kendini yenileyemez for döngüsü içinde olmadığı için; ondan dolayı böyle işe yaradı kod
            {
                if (checkedListBox1.GetItemCheckState(i) == CheckState.Checked)
                {
                    x++;
                    metin = x + " adet dosya gönderilmek üzere ek kısmına eklenmiştir.";

                    dosyasayikontrol.Text = metin;
                }
                if(checkedListBox1.GetItemCheckState(i) == CheckState.Unchecked)
                {
                    listBox5.Items.RemoveAt(i);
                    checkedListBox2.Items.RemoveAt(i);
                    checkedListBox1.Items.RemoveAt(i);
                    //tamam artık hem boyut hem de normal olarak silinebiliyorlar
                    Mbcheck();
                }
            }
            if (checkedListBox1.CheckedItems.Count == 0)
            {
                metin = "Şu an için yüklü bir dosya bulunmamaktadır";

                dosyasayikontrol.Text = metin;
            }

            txt_taslak.Text = txt_font_baslik.Text + txt_baslik.Text + Environment.NewLine + Environment.NewLine + txt_font_hitap.Text + txt_hitap.Text + " (...), " + Environment.NewLine + Environment.NewLine + txt_font_icerik.Text + txt_icerik.Text + Environment.NewLine + Environment.NewLine + txt_ekdosya.Text + dosyasayikontrol.Text;
        } //Ç
        void Check2()
        {
            string[] bakalim = new string[checkedListBox1.Items.Count];
            for (int f = 0; f < checkedListBox1.Items.Count; f++)
            {
                bakalim[f] = checkedListBox1.Items[f].ToString();
            }
            for (int k = 0; k < checkedListBox1.Items.Count; k++)
            {
                for (int l = 0; l < checkedListBox1.Items.Count; l++)
                {
                    if (bakalim[k] == bakalim[l])
                    {
                        if (k != l) //
                        {
                            MessageBox.Show("Lütfen daha önceden eklemiş olduğunuz dosyaları tekrardan eklemeyiniz. Aynı dosyayı birden fazla kez ekleme girişiminiz program tarafından silinmiştir. Aynısını tekrardan eklemeye çalıştığınız mevcut dosyanın ismi: " + bakalim[k].ToString(), "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            string[] kontrol3 = new string[checkedListBox1.Items.Count];
                            string[] kontrol4 = new string[checkedListBox2.Items.Count];
                            int[] değer2 = new int[listBox5.Items.Count];
                            for (int h = 0; h < checkedListBox1.Items.Count; h++)
                            {
                                kontrol3[h] = checkedListBox1.Items[h].ToString();
                                kontrol4[h] = checkedListBox2.Items[h].ToString();
                                değer2[h] = Convert.ToInt32(listBox5.Items[h].ToString());
                            }

                            string[] kontrol = kontrol3.Distinct().ToArray();
                            string[] kontrol2 = kontrol4.Distinct().ToArray();
                            int[] değer = değer2.Distinct().ToArray();

                            checkedListBox1.Items.Clear();
                            checkedListBox2.Items.Clear();
                            listBox5.Items.Clear();

                            for (int g = 0; g < kontrol.Length; g++)
                            {
                                checkedListBox1.Items.Add(kontrol[g]);
                                checkedListBox2.Items.Add(kontrol2[g]);
                                listBox5.Items.Add(değer[g]);
                            }

                            Mbcheck();
                            //şimdi, aynı olanı tespit etmenin yolu isimleri kontrol etmekten geçer ama mantıken burada k küçük olan l büyük olan olacak ister istemez çünkü son eklenen üsttekilerle aynı olabilir sadece geri kalanların farklı olması onaylanmış durum zaten ve dosyadan seçilenin boyutu ve ismi her zaman listboxlara bir kez de olsa ekleniyor yani boyut son değerde tutulmuş oluyor
                            //yani en son eklenen boyutu eğer çıkaracak olursam o zaman ilgili dosyanın boyutunu da ikinci kez hesaplamaktan kurtulmuş olacağım demektir bu
                            //bunun sayesinde eğer tekrar eklenmeye çalışılan olursa onun kontrolünü yapıyoruz ve aynıysa hem isim olarak hem de boyut olarak silmiş oluyoruz

                            //hmm üstteki açıklama ile yaptığım kod sağlıksız oldu; ona chechedlistbolara davrandığım gibi davranınca mükemmel sonucu alabileceğimi fark ettim ve onu uygulayaınca sorunum çok güzel çözülmü oldu, şu an maşallah var tıkır tıkır çalışıyor
                        }
                    }
                }
            }

            for (int b = 0; b < checkedListBox1.Items.Count; b++)
            {
                checkedListBox1.SetItemChecked(b, true);
            }
        } //D
        void Hataayiklama()
        {
            OleDbCommand komut;
            OleDbCommand komut2;
            string vtyolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=veri.accdb;Persist Security Info=True";
            OleDbConnection baglanti = new OleDbConnection(vtyolu);

            for (int a = 0; a < metroGrid1.Rows.Count; a++)
            {
                string isim = metroGrid1.Rows[a].Cells[0].Value.ToString();
                string e_mail = metroGrid1.Rows[a].Cells[1].Value.ToString();


                baglanti.Open();
                string ekle = "insert into veri(isim,mail) values (@isim,@mail)";
                komut = new OleDbCommand(ekle, baglanti);
                komut.Parameters.AddWithValue("@isim", isim.ToString());
                komut.Parameters.AddWithValue("@mail", e_mail.ToString());
                komut.ExecuteNonQuery();
                komut.Dispose();
                baglanti.Close();
                //tamam şu an veriler accesse eklendi artık, silinmesi gerekenler hala orada duruyor
            }

            for (int b = 0; b < listBox3.Items.Count; b++)
            {
                string h = listBox3.Items[b].ToString();

                baglanti.Open();
                string sil = "delete from veri where mail=@mail";
                komut2 = new OleDbCommand(sil, baglanti);
                komut2.Parameters.AddWithValue("@mail", h);
                komut2.ExecuteNonQuery();
                komut2.Dispose();
                baglanti.Close();
                //tamam şu an veriler silindi 
            }

            string pathconn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=veri.accdb;Persist Security Info=True";
            OleDbConnection conn = new OleDbConnection(pathconn);
            OleDbDataAdapter MyDataAdapter = new OleDbDataAdapter("Select * from veri", conn);
            DataTable dt = new DataTable();
            MyDataAdapter.Fill(dt);
            metroGrid1.DataSource = dt;
            Tasarim_2();
            metroLabel11.Text = metroGrid1.Rows.Count.ToString();

            metroGrid1.ClearSelection();

            baglanti.Open();
            string sil2 = "delete from veri";
            komut2 = new OleDbCommand(sil2, baglanti);
            komut2.ExecuteNonQuery();
            komut2.Dispose();
            baglanti.Close();

            int x = listBox3.Items.Count;

            if (listBox3.Items.Count == 0)
            {
                listBox3.Items.Add("Hatalı mail adresi bulunamadı");
            }
            else if(listBox3.Items.Count == 1)
            {
                listBox3.Items.Add("-----------------------");
                listBox3.Items.Add("1 adet hatalı mail adresi çıkarıldı");
            }
            else
            {
                listBox3.Items.Add("-----------------------");
                listBox3.Items.Add(x.ToString() + " adet hatalı mail adresleri çıkarıldı");
            }

            metroLabel11.Visible = true;

            if(metroLabel11.Text == "0")
            {
                metroTile4.Enabled = false;
                metroTile2.Enabled = false;
                MessageBox.Show("Yüklemiş olduğunuz excelde herhangi bir veri bulunamadı ya da veriler içinden geçerli olan herhangi bir mail adresi bulunamadı. Lütfen excel dosyanızın içinde verilerin olduğundan ve var olan mail adreslerinin doğru olduğundan emin olunuz ve ardından tekrar deneyiniz.","Hata",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            else if(metroLabel11.Text != "0")
            {
                metroTile4.Enabled = true;
                metroTile2.Enabled = true;
            }
        } //E
        void Satir()
        {
            listBox1.Items.Clear();
            string[] satirlar = new string[txt_icerik.Lines.Count()];
            for (int i = 0; i < txt_icerik.Lines.Count(); i++)
            {
                satirlar[i] = txt_icerik.Lines[i].ToString() + "<br>";
                listBox1.Items.Add(satirlar[i]);
            }
        } //F
        void Tasarim_2()
        {
            metroGrid1.Columns[0].HeaderText = "İsim ve Soyisim";
            metroGrid1.Columns[1].HeaderText = "E-mail";
            metroGrid1.Columns[0].Width = 137;
            metroGrid1.Columns[1].Width = 180;
        } //G
        void clbcheck()
        {
            if (checkedListBox1.Items.Count == 0)
            {
                checkedListBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            }
            else
            {
                checkedListBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            txt_taslak.Text = txt_font_baslik.Text + txt_baslik.Text + Environment.NewLine + Environment.NewLine + txt_font_hitap.Text + txt_hitap.Text + " (...), " + Environment.NewLine + Environment.NewLine + txt_font_icerik.Text + txt_icerik.Text + Environment.NewLine + Environment.NewLine + txt_ekdosya.Text + dosyasayikontrol.Text;

            txt_huptgmail.Text = Settings.Default.mail;
            txt_sifre.Text = Settings.Default.sifre;
        }  //Ğ
        private void MetroTile1_Click(object sender, EventArgs e)
        {
            try
            {
                listBox3.Items.Clear();
                OpenFileDialog openfile1 = new OpenFileDialog
                {
                    Filter = "Excel Dosyası |*.xlsx| Excel Dosyası|*.xls",
                    Title = "Veri Excel'ini seçiniz..."
                };
                if (openfile1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    this.textBox1.Text = openfile1.FileName;
                }

                Excel.Application oXL = new Excel.Application(); //hmm demek nuget paketten bulmak gerekiyormuş seni ve sonrada öyle using Excel diyerek kullanmak gerekiyormuş
                if (textBox1.Text == string.Empty)
                {
                    return;
                }
                else
                {
                    Excel.Workbook oWB = oXL.Workbooks.Open(textBox1.Text); // hata burada oluşuyor demek

                    List<string> liste = new List<string>();
                    foreach (Excel.Worksheet oSheet in oWB.Worksheets)
                    {
                        liste.Add(oSheet.Name);
                    }
                    oWB.Close();
                    oXL.Quit();
                    oWB = null;
                    oXL = null;
                    metroGrid2.DataSource = liste.Select(x => new { SayfaAdi = x }).ToList();
                    textBox2.Text = metroGrid2.Rows[0].Cells[0].Value.ToString();

                    OleDbCommand komut = new OleDbCommand();
                    string pathconn = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + textBox1.Text + ";Extended Properties=\"Excel 8.0;HDR= yes;\";";
                    OleDbConnection conn = new OleDbConnection(pathconn);
                    OleDbDataAdapter MyDataAdapter = new OleDbDataAdapter("Select * from [" + textBox2.Text + "$]", conn);
                    DataTable dt = new DataTable();
                    MyDataAdapter.Fill(dt);
                    metroGrid1.DataSource = dt;
                    Tasarim_2();
                    metroLabel11.Text = metroGrid1.Rows.Count.ToString();

                    metroGrid1.ClearSelection();

                    Regex regex = new Regex(@"^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$");

                    string[] email = new string[metroGrid1.Rows.Count];
                    for (int i = 0; i < metroGrid1.Rows.Count; i++)
                    {
                        email[i] = metroGrid1.Rows[i].Cells[1].Value.ToString();
                        string y = email[i].Trim(); //iyi oldu bu ya güzel işe yarıyor
                        email[i] = y.ToString();

                        bool İsValidEmail = regex.IsMatch(email[i]);

                        if (!İsValidEmail)
                        {
                            listBox3.Items.Add(email[i].ToString());
                        }
                    }

                    if (metroLink3.Text == "Hatalı Mail Adreslerini Gizle")
                    {
                        listBox3.Visible = true;
                    }
                    else if (metroLink3.Text == "Hatalı Mail Adreslerini Göster")
                    {
                        listBox3.Visible = false;
                    }

                    metroGrid1.BorderStyle = System.Windows.Forms.BorderStyle.None;
                }

                Hataayiklama();
            }
            catch(Exception)
            {
                MessageBox.Show("Yüklemeye çalıştığınız excel dosyasını lütfen kapatınız. Açık olan excel dosyanızı kapattıktan sonra programa veri girişini yapmayı lütfen tekrar deneyiniz. \n\nEk Bilgi: Hem herhangi bir programda hem de bilgisayarda eş zamanlı olarak aynı excel dosyası çalıştırılamaz.", "Hata",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }
        } //I
        private void MetroTile4_Click(object sender, EventArgs e)
        {  
            OpenFileDialog myFileDialog = new OpenFileDialog();
            DialogResult dr;
            myFileDialog.Title = "Dosya/Dosyalar ekle...";
            myFileDialog.InitialDirectory = @"C:";
            myFileDialog.Multiselect = true;
            dr = myFileDialog.ShowDialog();
            string[] fileNames = myFileDialog.FileNames;

            long[] boyut = new long[fileNames.Length];
            for (int i = 0; i < fileNames.Length; i++)
            {
                FileInfo bilgi = new FileInfo(fileNames[i].ToString());
                long boyut2 = bilgi.Length;
                boyut[i] = boyut2;

                listBox5.Items.Add(boyut[i]);
            }

            myFileDialog.CheckFileExists = true;
            myFileDialog.CheckPathExists = true;

            long sinir = 5000000;

            if (dr == DialogResult.OK)
            {
                long toplam = 0;
                long[] dizi2 = new long[listBox5.Items.Count];
                for (int i = 0; i < listBox5.Items.Count; i++)
                {
                    toplam += Convert.ToInt32(listBox5.Items[i].ToString());
                    dizi2[i] = Convert.ToInt32(listBox5.Items[i].ToString());
                }

                if (toplam >= sinir)
                {
                    listBox5.Items.Clear();
                    for (int i = 0; i < dizi2.Length - 1; i++)
                    {
                        listBox5.Items.Add(dizi2[i].ToString());
                    }

                    MessageBox.Show("Toplam dosya boyutu 5 MB'ı aşıyor. Lütfen seçtiğiniz dosyaların toplam 5 MB'ı aşmadığından emin olunuz.", "Dosya Boyutu Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    clbcheck();
                }
                else
                {
                    string[] dosya = fileNames;

                    for (int i = 0; i < fileNames.Length; i++)
                    {
                        checkedListBox1.Items.Add(dosya[i]);
                        checkedListBox2.Items.Add(dosya[i]);
                    }

                    string[] isimler = new string[checkedListBox1.Items.Count];
                    string[] dosyaadi = new string[checkedListBox1.Items.Count];
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        isimler[i] = checkedListBox2.Items[i].ToString();
                        string m = TersCevir(isimler[i]);
                        int dur = m.IndexOf(@"\");
                        string ters = m.Substring(0, dur);
                        string düzelt = TersCevir(ters);
                        dosyaadi[i] = düzelt;
                    }

                    checkedListBox1.Items.Clear();
                    checkedListBox2.Items.Clear();

                    for (int i = 0; i < dosyaadi.Length; i++)
                    {
                        checkedListBox1.Items.Add(dosyaadi[i]);
                        checkedListBox2.Items.Add(isimler[i]);
                    }

                    clbcheck();
                }
            }

            for (int b = 0; b < checkedListBox1.Items.Count; b++)
            {
                checkedListBox1.SetItemChecked(b, true);
            }

            Check1();
            Check2();
            Mbcheck();


        } //İ
        private void CheckedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Check1();
        } //J
        private void Txt_icerik_TextChanged(object sender, EventArgs e)
        {
            Satir();
            txt_icerik2.Text = "";
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                txt_icerik2.Text += listBox1.Items[i].ToString(); //hmm, bileşik atama da güzelmiş
            }

            txt_taslak.Text = txt_font_baslik.Text + txt_baslik.Text + Environment.NewLine + Environment.NewLine + txt_font_hitap.Text + txt_hitap.Text + " (...), " + Environment.NewLine + Environment.NewLine + txt_font_icerik.Text + txt_icerik.Text + Environment.NewLine + Environment.NewLine + txt_ekdosya.Text + dosyasayikontrol.Text;
        } //K
        private void Txt_baslik_TextChanged(object sender, EventArgs e)
        {
            txt_taslak.Text = txt_font_baslik.Text + txt_baslik.Text + Environment.NewLine + Environment.NewLine + txt_font_hitap.Text + txt_hitap.Text + " (...), " + Environment.NewLine + Environment.NewLine + txt_font_icerik.Text + txt_icerik.Text + Environment.NewLine + Environment.NewLine + txt_ekdosya.Text + dosyasayikontrol.Text;
        } //K
        private void Txt_hitap_TextChanged(object sender, EventArgs e)
        {
            txt_taslak.Text = txt_font_baslik.Text + txt_baslik.Text + Environment.NewLine + Environment.NewLine + txt_font_hitap.Text + txt_hitap.Text + " (...), " + Environment.NewLine + Environment.NewLine + txt_font_icerik.Text + txt_icerik.Text + Environment.NewLine + Environment.NewLine + txt_ekdosya.Text + dosyasayikontrol.Text;
        } //K
        private void MetroTile2_Click(object sender, EventArgs e)
        {
            if(txt_icerik.Text == "" | txt_hitap.Text == "" || txt_baslik.Text == "")
            {
                string[] dizi = { txt_icerik.Text, txt_hitap.Text, txt_baslik.Text};
                for (int i = 0; i < 3; i++)
                {
                    if (dizi[i] == "")
                    {
                        if (i == 0)
                        {
                            txt_icerik.WaterMark = "Mail İçeriğini Giriniz... (Lütfen doldurunuz.)";
                        }
                        else if (i == 1)
                        {
                            txt_hitap.WaterMark = "Hitap Şeklinizi Giriniz...(Ör: Merhaba Sevgili) (Lütfen doldurunuz.)";
                        }
                        else if (i == 2)
                        {
                            txt_baslik.WaterMark = "Mail Başlığını Giriniz... (Lütfen doldurunuz.)";
                        }
                    }
                }

                MessageBox.Show("Lütfen tüm kutucukları doldurunuz! Doldurmanız gereken kutucukları kutucukların üstündeki yazıdan görebilirsiniz.","Hata",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
            else
            {
                DialogResult dialog;
                dialog = MessageBox.Show("Katılımcılara toplu bir şekilde e-mail göndermek istediğinize emin misiniz?", "Mail Taslağı Kontrolü", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialog == DialogResult.Yes)
                {
                    metroLabel8.Visible = true;
                    metroProgressBar1.Visible = true;
                    metroLabel9.Visible = true;

                    for (int i = 0; i <= metroGrid1.Rows.Count; i++)
                    {
                        //***0 dan başlıyor for döngüsü işlemeye ve datagridin başlıkları 0.eleman olarak kabul edilmiyor ilk değeri 0.eleman olarak kabul ediliyor***
                        //***sorun 0.ıncı değerde hiç ilerlenilmemesi***
                        if (i == metroGrid1.Rows.Count)
                        {
                            decimal yuzde = ((decimal)(i) / (decimal)metroGrid1.Rows.Count) * 100;
                            Application.DoEvents();
                            metroProgressBar1.Value = (int)yuzde;
                            yuzde = Math.Round(yuzde, 2);
                            metroLabel8.Text = "%" + yuzde.ToString();
                        }
                        else
                        {
                            txt_isim.Text = metroGrid1.Rows[i].Cells[0].Value.ToString();
                            txt_email.Text = metroGrid1.Rows[i].Cells[1].Value.ToString();
                            Emailgonder();

                            decimal yuzde = ((decimal)(i + 1) / (decimal)metroGrid1.Rows.Count) * 100;
                            Application.DoEvents();
                            metroProgressBar1.Value = (int)yuzde;
                            yuzde = Math.Round(yuzde, 2);
                            metroLabel8.Text = "%" + yuzde.ToString();
                        }
                    }

                    metroLabel9.Visible = false;

                    Form_Loading x = new Form_Loading();
                    x.Show();
                }

                else
                {
                    MessageBox.Show("e-mail gönderme işlemi başlatılmadı.", "Başlatılmayan İşlem", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        } //L
        private void MetroLink3_Click(object sender, EventArgs e)
        {
            if (lst3_kutucuk.Items.Count == 0)
            {
                int butonabasissayisi = 0;
                butonabasissayisi++;
                lst3_kutucuk.Items.Add(butonabasissayisi.ToString());

                if (butonabasissayisi % 2 == 0)
                {
                    metroLink3.Text = "Hatalı Mail Adreslerini Göster";
                    listBox3.Visible = false;
                }
                else
                {
                    metroLink3.Text = "Hatalı Mail Adreslerini Gizle";
                    listBox3.Visible = true;
                    if (listBox3.Items.Count == 0)
                    {
                        listBox3.Items.Add("Hatalı mail adresleri bulunamadı");
                    }
                }
            }
            else
            {
                int yeni = Convert.ToInt32(lst3_kutucuk.Items.Count.ToString());
                yeni++;
                lst3_kutucuk.Items.Add(yeni.ToString());

                if (yeni % 2 == 0)
                {
                    metroLink3.Text = "Hatalı Mail Adreslerini Göster";
                    listBox3.Visible = false;
                }
                else
                {
                    metroLink3.Text = "Hatalı Mail Adreslerini Gizle";
                    listBox3.Visible = true;
                    if (listBox3.Items.Count == 0)
                    {
                        listBox3.Items.Add("Hatalı mail adresleri bulunamadı");
                    }
                }
            }
        } //M
        private void MetroLink2_Click(object sender, EventArgs e)
        {
            Form3 x = new Form3();
            x.Show();
        } //N
        private void metroLink1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://myaccount.google.com/lesssecureapps");
        }
        private void checkedListBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            clbcheck();
        }
        #region
        //A
        /*
         A kısmında dosya yolları alınan metinleri tersten alma ihtiyacı hissettim çünkü insanlar dosya isimlerini uzun uzun yollarıyla değil normal şekillerde görsünler
        diye bu metotu kullandım. Kendim tasarlamadım internetten hazır aldım
         */

        //B
        /*
         B kısmı klasik e mail gönderme kodları ufak tefek eklemelerle; toplu bir şekilde kişilere özel mailler attığımız için bir defa bir void yapısının içinde olması kod yazımı
        için çok daha rahat olacaktır, ekstra olarak dosya gönderimleri için alınan dosya sayısı kadar döndürme yaparak tek mailde birden fazla ek yollayabilme imkanı oluşturmuş
        olduk; ayrıca, timeout özelliğini yazmamız şart çünkü bazen mail gönderimleri eklerin boyutuna bağlı olarak ya da internetin yavaş olabilme durumundan kaynaklı olarak
        uzun sürebilme ihtimali olduğu için bunu 30 dk ya da 300 dk olarak ayarladım sıfırlardan tam emin olamadım ama en az 30 dk her bir mail için garanti ayrılan süre var ki bu süre
        tost makinesinden internete bağlanabilen bir şey için bile yeterli bir süre dsgadsgd

        "Hata Ayıkla > Pencereler > Özel Durum Ayarları > Managed Debugging Assistants" kısmından "ContextSwitchDeadlock" ayarını kapatmayı asla unutmayın çünkü o ayar sayesinde kafamıza
        göre timeout ayarlayabiliyoruz, yoksa gmail bizim göndermeye çalıştığımız maili zaman aşımına uğramış kabul edip hata veriyor kodda hiç bir hata olmamasına rağmen aman dikkat
         */

        //C
        /*
         C kısmı baytı mb'a dönüştürmeye yarayan klasik kodumuz. Dosyaları alırken biz bayt cinsinden alırız e bayt olarak programda kullanıcıya direkt göstermek hoş olmaz
        ondan dolayı bu dönüştürücüye ihtiyaç duydum
         */

        //Ç
        /*
         Ç kısmı tamamen taslak metnimizin daha eşzamanlı gözükmesi adına koydum. Eklediğin çıkardığın her dosya bu kısım sayesinde taslak kısmında ekte şu kadar dosya mevcut
        şeklinde yansıtmaya yarıyor
         */

        //D
        /*
         D kısmı ise aynı eğer eklenirse yanlışlıkla eklerde 2 kere gitmesin diye var. Aynı dosyayı yüklediysen algılıyor ve onu eklere dahil edemeyeceğini çünkü zaten halihazırda eklerde
        bu dosyanın var olduğunu söylüyor
         */

        //E
        /*
         E kısmında Hatalı Mail Adreslerini çıkarmak için veritabanlarından faydalandım; ham halde excelden gride aktarılan her türlü mail burada kontrol edilerek rafine bir şekilde
        geri aynı gridde gözükmesini sağlıyor ve üstelik hatalı maillerin ne olduğunu hangi maillerin çıkartıldığını görüyorsun bu sayede çok çok çok büyük ölçüde mail gönderimi sırasında
        mail formatına aykırı olan maillerin mail gönderimi baltalamasına programın hata vermesinin önüne geçiyorsun

        Programın kilit kısımlarından biridir
         */

        //F
        /*
         F kısmı da programın kilit kısımlarından biridir. Normalde mail gönderirken sen mesajını html textboxlarına yazarsın ve onlar sayfada otomatik olarak satır atlatmayı <br> kodunu
        göstermeden yapabilir ama c# da html textbox olmadığı için program üstünden mail gönderimi yaparken nerede boşluk koyacaksın bunu c# otomatik olarak yapmıyor; bu yüzden enterla alta geçilen
        her satırı bir diziye aktardım ve diziye aktarırken de <br> takısını ekledim bu sayede tıpkı sanki taslak hazırlarken olduğu gibi o <br> taglarını sen yazarken görmüyorsun ama program otomatik olarak
        arkaplanda her satırın sonuna o br takısını ekliyor ve bu sayede mailimiz tam olarak taslak metinde nasıl görüyorsak aynı şekilde aktarılmış oluyor
         */

        //G
        /*
         Anlatılacak bir şey yok ya sadece datagrid(ya da bu metroframework kapsamında metrogrid) sütun başlıkları düzgün gözüksün diye ekledim
         */

        //Ğ
        /*
         * Anlatılacak bir şey yok ya sadece taslak metin için gereken kutucukları birleştirdim
         */

        //H
        /*
         İki tane H bıraktım bilerek çünkü o ikisi birbiriyle gövdeden bağlı, çok basit bir şekilde şifrenin açılıp kapanmasını sağlıyor ama sonradan farkettim ki şifre gözükmemeli
        çünkü topluluğun şifresi gözükmüş olur o şekilde :D o yüzden burayı daha sonra kaldırmak zorunda kaldım ama silmeye de kıyamadım :D

        private void PictureBox1_Click(object sender, EventArgs e)
        {
            pictureBox1.Visible = false;
            pictureBox2.Visible = true;

            txt_sifre.PasswordChar = '*';
        }  //H
        private void PictureBox2_Click(object sender, EventArgs e)
        {
            pictureBox1.Visible = true;
            pictureBox2.Visible = false;

            txt_sifre.PasswordChar = '\0';
        } //H

         */

        //I
        /*
         Kilit Kısım. Veriler Gride aktarılır ve daha sonrada hatalı olanlar ayıklanır üstteki voidlerin yardımıyla
         */

        //İ
        /*
         Kilit Kısım. Dosyaların eklenmesi/çıkarılması 5 MB geçip geçmediği gibi vb. dosyalarla alakalı tüm kontroller burada üstteki voidlerin yardımıyla
         */

        //J
        /*
         Çok basit şekilde dosyayı çıkarma yöntemi olarak dosyanın üstüne basarsa çıkartırsın şeklinde düşündüm o kadar öteki türlü çok hızlı bakarsa tiksiz kalma ihtimali var
        ona ayrı kod yazmak yerine tek tıklama ile dosyayı kaldırmayı seçtim tıpkı gmail taslaklarında da tek tıkla silindiği gibi
         */

        //K
        /*
         üçü de aynı işlevde, taslak görünüşünü nasıl olur diye ayarlamalarını yaptığım kısımlar
         */

        //L
        /*
         Mail gönderme tuşu, mail gönderimi ile alakalı tüm kontroller burada üstteki voidlerin yardımıyla
         */

        //M
        /*
         Hatalı maillerin görünüp görünmeyeceğini belirten link, eğer görmez istersen tıklarsın eğer görmek istemiyorsan da kapatırsın
         */

        //N
        /*
         Program Kullanım Klavuzunu açan kısım
         */

        /*catch (SmtpFailedRecipientsException)
{
    return;
}
catch (SmtpFailedRecipientException)
{
    return;
}
catch (SmtpException)
{
    MessageBox.Show("Bu hatayı almanızın büyük ihtimalle üç sebebi vardır: \n\n1. Lütfen toplu mail gönderimini gerçekleştireceğiniz gmail hesabının" + @" ""Less Secure Apps"" " + "özelliğinin açık olduğundan emin olunuz(özelliğin linki programın içinde mevcuttur, oraya tıklamanızın ardından otomatik olarak yönlendirileceksiniz).\n(%90) \n\n2. Lütfen toplu mail gönderimini gerçekleştireceğiniz gmail hesabının mail adresini ve şifresini doğru girdiğinizden emin olunuz.\n(%8) \n\nBu ikisini kontrol etmenizin ve doğruluğundan emin olmanızın ardından lütfen mailleri göndermeyi tekrardan deneyiniz. \n\n3. Pek zannetmiyorum ama google ücretsiz mail servis sağlayıcısı desteğini filan kesmişse ya da ayarlarında değişikliğe gitmişse eğer program tost olmuş demektir, yeni bir ücretsiz mail servis sağlayıcısı ayarlamak gerekir programa eğer öyle bir şey söz konusu ise.\n(%2) \n\nBu üçünü kontrol etmenize ve herhangi bir sorun olmadığına emin olmanıza rağmen hala bu hatayı alıyorsanız bulunduğunuz yerde ufak çaplı bir teknolojik kıyamet filan yaşandı herhalde ve siz buna rağmen toplu mail gönderme derdine düştünüz, elektrik ve internet bağlantısı mevcutken 1 maili 1 saatte gönderemeyecek kadar tost değildir internetiniz diye düşünmek istiyorum.\n(≈%0,00....1) \n\nGözümden kaçan başka bir hata durumu da oluşmuş olabilir. Eğer sorun çözülemezse bana" + @" ""muhammettahirtekatli@gmail.com"" " + "mail adresinden lütfen ulaşın.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
}
catch (Exception)
{
    //hmm bak gmail bile öyle bir mail adresi yok diyor ama yine de gönderildi diyor(doğru, return dedik ya abi onları görürsen takılma diye e programda onları takmadı ve tüm maillerini gönderdiğini söyledi; sonuçta program da kendi içinde haklı çünkü bizim "gönder" dediğimiz tüm mail adreslerine o göndermiş oldu aslında, gönderme dediğimize göndermedi program da haliyle, aferin uslu çocuk :)))
    return;
}
//hmmm try catch hatalarını böyle böyle tanıtmak çok iyi ya, çok lazım oluyor onlar gerçekten çünkü çok işlevseller*/
        #endregion

        private void metroLink4_Click(object sender, EventArgs e)
        {
            sifre x = new sifre();
            x.Show();
        }
    }
}
