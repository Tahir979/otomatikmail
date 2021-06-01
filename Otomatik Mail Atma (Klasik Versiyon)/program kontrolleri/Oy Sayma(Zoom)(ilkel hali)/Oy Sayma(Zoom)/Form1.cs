using System;
using System.Windows.Forms;
using System.Net;
using System.Net.Mail;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Linq;

namespace Oy_Sayma_Zoom_
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        OleDbConnection con;
        OleDbDataAdapter da;
        DataSet ds;
        DataSet ds2;
        DataSet ds3;
        DataTable dt = new DataTable();
        void griddoldur()
        {
            con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=oy.accdb");
            da = new OleDbDataAdapter("SElect * from Oykod", con);
            ds2 = new DataSet();
            con.Open();
            da.Fill(ds2, "Oykod");
            dataGridView2.DataSource = ds2.Tables["Oykod"];
            con.Close();
        }
        void griddoldur2()
        {
            con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=oy.accdb");
            da = new OleDbDataAdapter("SElect * from kontrol", con);
            ds = new DataSet();
            con.Open();
            da.Fill(ds, "kontrol");
            dataGridView4.DataSource = ds.Tables["kontrol"];
            con.Close();
        }
        void griddoldur3()
        {
            con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=oy.accdb");
            da = new OleDbDataAdapter("SElect * from gecersiz", con);
            ds3 = new DataSet();
            con.Open();
            da.Fill(ds3, "gecersiz");
            dataGridView5.DataSource = ds3.Tables["gecersiz"];
            con.Close();
        }
        void datakaydet()
        {
            string vtyolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=oy.accdb";
            OleDbConnection baglanti = new OleDbConnection(vtyolu);
            baglanti.Open();
            string ekle = "insert into Oykod(İsim_table,Email_table,Kod_table) values (@İsim_table,@Email_table,@Kod_table)";
            OleDbCommand komut = new OleDbCommand(ekle, baglanti);
            komut.Parameters.AddWithValue("@İsim_table", txt_isim.Text);
            komut.Parameters.AddWithValue("@Email_table", txt_email.Text);
            komut.Parameters.AddWithValue("@Kod_table", txt_kod.Text);
            komut.ExecuteNonQuery();
        }
        void emailgonder()
        {
            try
            {
                SmtpClient sc = new SmtpClient();
                sc.Port = 587;
                sc.Host = "smtp.gmail.com";
                sc.EnableSsl = true;
                sc.Credentials = new NetworkCredential("hupsikoloji@gmail.com", "Yk2021adler");
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress("hupsikoloji@gmail.com", "HÜPT");
                mail.To.Add(txt_email.Text.ToString());                                                                                                                                                       
                mail.Subject = "HÜPT Etkinlik"; mail.IsBodyHtml = true; mail.Body = "Merhaba Sevgili " + txt_isim.Text.ToString() + "," + "<br/> <br/>" + "Bugün saat 17.00'de gerçekleştirilecek olan " + @" ""Etkili CV Hazırlama Yöntemleri"" " + " adlı etkinliğimize katılabilmeniz için gereken zoom bağlantı linki: https://us05web.zoom.us/j/3738806292?pwd=RFVtdmlYRFE3VTJ5dUE4ZUdtL0xKQT09 (Meeting ID: 373 880 6292 Şifre: hupt)" + "<br/> <br/>" + "Değerli katılımınıza çok teşekkür ederiz."; // + "<br/> <br/>" + "Bugünün ikinci etkinliği ve 1.Bilişsel Psikoloji Haftamızın son etkinliği olan saat 20.00'de gerçekleştireceğimiz" + @" ""Bir Dinamik Sistem Olarak Davranış ve Biliş"" " + " konulu seminere katılabilmeniz için gereken zoom bağlantı linki: https://zoom.us/j/92727603100 (Meeting ID: 927 2760 3100)." + "<br/> <br/>" + " 1.Bilişsel Psikoloji Haftası sonunda dijital katılım sertfikası alabilmek için bu hafta boyunca düzenlenecek olan 5 seminerimizin 4'üne katılımın zorunlu olduğunu belirterek değerli katılımınıza çok teşekkür ederiz. Ayrıca seminere olan katılımınızı bize belirtebilmek adına lütfen zoom chatten gönderilecek olan formu doldurmayı unutmayalım.";
                //mail.Subject = "Hüpt Pusula"; mail.IsBodyHtml = true; mail.Body = "Merhaba Sevgili " + txt_isim.Text.ToString() + "," + "<br/> <br/>" + "Yaşadığımız bir teknik aksaklık sonucu etkinlik linkimizi değiştirmiş bulunmaktayız, bu son dakika değişikliği için özür dileriz. Bugün saat 17.00'de " + @" ""Sosyal Psikolojiye Alandan ve Akademiden Bakış"" " + " adlı etkinliğimizin " + @" ""Alandan Bakış"" " + " kısmını gerçekleştireceğiz. Bu etkinliğimize katılabilmeniz için gereken zoom bağlantı linki: https://us05web.zoom.us/j/3738806292?pwd=RFVtdmlYRFE3VTJ5dUE4ZUdtL0xKQT09 (Meeting ID: 373 880 6292 Şifre: hupt)." + "<br/> <br/>" + "Değerli katılımınıza çok teşekkür ederiz.";
                sc.Send(mail);
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString() + "Oy gönderiminde bir hata ile karşılaşıldı");
            }
        }
        void kod()
        {
            string GuvenlikKodu;
            GuvenlikKodu = "";
            int harf, bykharf, hangisi;
            Random Rharf = new Random();
            Random Rsayi = new Random();
            Random Rbykharf = new Random();
            Random Rhangisi = new Random();

            for (int b = 0; b < 6; b++)
            {
                int a = 0;
                hangisi = Rhangisi.Next(1, 3);
                if (hangisi == 1)
                {
                    GuvenlikKodu += Rsayi.Next(0, 10).ToString();
                }
                if (hangisi == 2)
                {
                    harf = Rharf.Next(1, 27);
                    for (char i = 'a'; i <= 'z'; i++)
                    {
                        a++;
                        if (a == harf)
                        {
                            bykharf = Rbykharf.Next(1, 3);
                            if (bykharf == 1)
                            {
                                GuvenlikKodu += i;
                            }
                            if (bykharf == 2)
                            {
                                GuvenlikKodu += i.ToString().ToUpper();
                            }
                        }
                    }
                }

            }
            txt_kod.Text = GuvenlikKodu;
        }
        void tasarim()
        {
            dataGridView2.Columns[0].HeaderText = "İsim ve Soyisim";
            dataGridView2.Columns[1].HeaderText = "E-mail";
            dataGridView2.Columns[2].HeaderText = "Doğrulama Kodu";
            dataGridView2.Columns[0].Width = 150;
            dataGridView2.Columns[1].Width = 180;
            dataGridView2.Columns[2].Width = 110;
        }
        void tasarim_2()
        {
            dataGridView1.Columns[0].HeaderText = "İsim ve Soyisim";
            dataGridView1.Columns[1].HeaderText = "E-mail";
            dataGridView1.Columns[0].Width = 150;
            dataGridView1.Columns[1].Width = 180;
        }
        void tasarim_3()
        {
            dataGridView3.Columns[0].HeaderText = "İsim ve Soyisim";
            dataGridView3.Columns[0].Width = 150;
            dataGridView3.Columns[1].HeaderText = "Doğrulama Kodu";
            dataGridView3.Columns[1].Width = 60;
            dataGridView3.Columns[2].HeaderText = "Başkanlık Oyları";
            dataGridView3.Columns[2].Width = 100;
            dataGridView3.Columns[3].HeaderText = "Bağımsız Aday Oyları";
            dataGridView3.Columns[3].Width = 200;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            griddoldur();
            griddoldur2();
            kod();
            tasarim();
            label10.Text = "?";
            label13.Text = "?";
            dataGridView4.Columns[0].HeaderText = "Bilgilendirme";
            dataGridView4.Columns[0].Width = 400;
            dataGridView4.Columns[1].Visible = false;


        }
        private void btn_kaydet_Click(object sender, EventArgs e)
        {
            datakaydet();
            emailgonder();
            kod();
            griddoldur();
            MessageBox.Show("İşlem Başarılı", "Bilgilendirme", MessageBoxButtons.OK, MessageBoxIcon.Information);
            txt_isim.Focus();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfile1 = new OpenFileDialog();
            if (openfile1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.textBox1.Text = openfile1.FileName;
            }

            string pathconn = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + textBox1.Text + ";Extended Properties=\"Excel 8.0;HDR= yes;\";";
            OleDbConnection conn = new OleDbConnection(pathconn);
            OleDbDataAdapter MyDataAdapter = new OleDbDataAdapter("Select * from [" + textBox2.Text + "$]", conn);
            DataTable dt = new DataTable();
            MyDataAdapter.Fill(dt);
            dataGridView1.DataSource = dt;
            tasarim_2();
            label10.Text = dataGridView1.Rows.Count.ToString();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            int kayitsayisi;
            kayitsayisi = dataGridView1.Rows.Count;
            listBox1.Items.Add("------İşlem Başlatılıyor...-------");
            for (int i = 0; i < kayitsayisi; i++)
            {
                txt_isim.Text = dataGridView1.Rows[i].Cells[0].Value.ToString();
                txt_email.Text = dataGridView1.Rows[i].Cells[1].Value.ToString();

                datakaydet();
                emailgonder();
                //kod();
                griddoldur();

                listBox1.Items.Add(txt_isim.Text + ", e-mail gönderme işlemi başarılı.");

                //int kayit;
                //kayit = dataGridView1.Rows.Count;

                //progressBar1.Minimum = 0;
                //progressBar1.Maximum = 100;
                //progressBar1.Value += 100 / kayit;
            }

            listBox1.Items.Add("------İşlem Tamamlandı...-------");
            listBox1.Items.Add(kayitsayisi.ToString() + " adet kişiye e-mail gönderildi");

            label13.Text = dataGridView2.Rows.Count.ToString();

            MessageBox.Show("İşlem tamamladı");
        }
        private void button6_Click(object sender, EventArgs e)
        {
            con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=oy.accdb"); 
            con.Open();
            OleDbCommand cmd = new OleDbCommand("SELECT Email_table FROM Oykod", con);
            OleDbDataReader dr = cmd.ExecuteReader();

            List<string> emailler = new List<string>();
            while (dr.Read())
            {
                emailler.Add(dr["Email_table"].ToString());
            }
            con.Close();
            string[] str = emailler.ToArray();

            int sayac = 0;

            for (int j = 0; j <= str.Length -1; j++)
            {
                for (int i = 0; i <= str.Length - 1; i++)
                {
                    if (str[j] == str[i])
                    {
                        for (int l = 0; l < j; l++)
                        {
                            if (str[l] == str[j])
                            {
                                sayac = -1;
                            }
                        }
                        sayac++;
                    }
                }
                if (sayac != 0)
                {
                    textBox6.Text = str[j] + " e-mail adresine " + sayac + " adet doğrulama kodu gönderildi.";

                    string vtyolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=oy.accdb";
                    OleDbConnection baglanti = new OleDbConnection(vtyolu);
                    baglanti.Open();
                    string ekle = "insert into kontrol(sonuclar) values (@sonuclar)";
                    OleDbCommand komut = new OleDbCommand(ekle, baglanti);
                    komut.Parameters.AddWithValue("@sonuclar", textBox6.Text);
                    komut.ExecuteNonQuery();
                    sayac = 0;
                }
            }
            MessageBox.Show("İşlem Tamamlandı");
            griddoldur2();
        }
        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            DataView dv = ds.Tables[0].DefaultView;
            dv.RowFilter = "sonuclar Like '%" + textBox5.Text + "%' or karisik Like '%" + textBox5.Text + "%'";
            dataGridView4.DataSource = dv;
        }
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                if (listBox1.Items[i].ToString().ToLower().Contains(textBox5.Text.ToLower()))
                {
                    listBox1.SetSelected(i, true);
                }
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfile1 = new OpenFileDialog();
            if (openfile1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.textBox1.Text = openfile1.FileName;
            }

            string pathconn = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + textBox1.Text + ";Extended Properties=\"Excel 8.0;HDR= yes;\";";
            OleDbConnection conn = new OleDbConnection(pathconn);
            OleDbDataAdapter MyDataAdapter = new OleDbDataAdapter("Select * from [" + textBox2.Text + "$]", conn);
            DataTable dt = new DataTable();
            MyDataAdapter.Fill(dt);
            dataGridView3.DataSource = dt;
            tasarim_3();
            label16.Text = dataGridView3.Rows.Count.ToString();


        }
        private void button2_Click_1(object sender, EventArgs e)
        {
            //Gönderdiğimiz Kodların Dizisi
            con = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=oy.accdb");
            con.Open();
            OleDbCommand cmd = new OleDbCommand("SELECT Kod_table FROM Oykod", con);
            OleDbDataReader dr = cmd.ExecuteReader();

            List<string> kod_1 = new List<string>();
            while (dr.Read())
            {
                kod_1.Add(dr["Kod_table"].ToString());
            }
            con.Close();
            string[] kodlar_1 = kod_1.ToArray(); //1.dizimiz
            Array.Sort(kodlar_1);

            for (int i = 0; i < kodlar_1.Length; i++)
            {
                listBox3.Items.Add(kodlar_1[i]);
            }

            //Aldığımız Kodların Dizisi
            List<string> kod_2 = new List<string>();
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                kod_2.Add(dataGridView3.Rows[i].Cells["doğrulama kodu"].Value.ToString());
            }

            string[] kodlar_2 = kod_2.ToArray(); //2.dizimiz
            Array.Sort(kodlar_2);

            for (int i = 0; i < kodlar_2.Length; i++)
            {
                listBox4.Items.Add(kodlar_2[i]);
            }

            //Dizileri Karşılaştırma
            var C = kodlar_2.Except(kodlar_1).ToList(); //bu sistem bu doğrulama kodunu göndermemiş, yani uydurulmuş bir doğrulama kodu bu, bu doğrulama kodunu yazan kişinin oyu sayılmayacaktır
            //BÜYÜK KÜÇÜK HARF UYUMUNU KONTROL ET ONDAN DA SORUN ÇIKIYOR MU DİYE
            //evet gösteriyor ama o sorun şöyle halledilecek, hani e mail ara dediğim yer var ya oraya kodu girdiğim zaman tamamen alakasız bir kod mu yoksa kişinin ufak çaplı harf hatasından mı kaynaklı bir sorun olduğunu tespit edebiliriz

            string[] kontrol = C.ToArray();
            Array.Sort(kontrol);

            for (int i = 0; i < kontrol.Length; i++)
            {
                listBox6.Items.Add(kontrol[i]);
            }
            if(listBox6.Items.Count == 0)
            {
                listBox6.Items.Add("Sistemin gönderdiği kodlardan farklı olan herhangi bir kod alınmamıştır.");
            }
            MessageBox.Show(listBox6.Items.Count.ToString());

            //Bu kodu silmek pek içimden gelmedi açıkcası çünkü beğendim sadece şu durum için yetersiz işleve sahip, yoksa sahip olduğu işlevi baya bi beğendim
            /*bool esitmi = kodlar_1.SequenceEqual(kodlar_2);
            MessageBox.Show(esitmi.ToString());*/

            //int x = listBox3.Items.Count;
            //int y = listBox4.Items.Count;
            //label28.Text = x.ToString();
            //label30.Text = y.ToString();
        }
        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            DataView dv = ds2.Tables[0].DefaultView;
            dv.RowFilter = "Kod_table Like '%" + textBox4.Text + "%' or Email_table Like '%" + textBox4.Text + "%'";
            dataGridView2.DataSource = dv;
        }
        private void button7_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                txt_isimvesoyisim.Text = dataGridView3.Rows[i].Cells[0].Value.ToString();
                txt_doğrulamakodu.Text = dataGridView3.Rows[i].Cells[1].Value.ToString();
                txt_baskanlikoylari.Text = dataGridView3.Rows[i].Cells[2].Value.ToString();
                txt_bagimsizoylari.Text = dataGridView3.Rows[i].Cells[3].Value.ToString();

                string vtyolu = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=oy.accdb";
                OleDbConnection baglanti = new OleDbConnection(vtyolu);
                baglanti.Open();
                string ekle = "insert into gecerlioylar(isim,kod,baskanlik,bagimsiz) values (@isim,@kod,@baskanlik,@bagimsiz)";
                OleDbCommand komut = new OleDbCommand(ekle, baglanti);
                komut.Parameters.AddWithValue("@isim", txt_isimvesoyisim.Text);
                komut.Parameters.AddWithValue("@kod", txt_doğrulamakodu.Text);
                komut.Parameters.AddWithValue("@baskanlik", txt_baskanlikoylari.Text);
                komut.Parameters.AddWithValue("@bagimsiz", txt_bagimsizoylari.Text);
                komut.ExecuteNonQuery();
            }
            MessageBox.Show("İşlem Tamamlandı");
        }
        private void dataGridView2_MouseClick(object sender, MouseEventArgs e)
        {
            textBox7.Text = dataGridView2.CurrentRow.Cells[0].Value.ToString();
            textBox12.Text = dataGridView2.CurrentRow.Cells[1].Value.ToString();

            DialogResult cikis = new DialogResult();
            cikis = MessageBox.Show("Tıkladığınız kişiyi silmek istediğinize emin misiniz? ?", "Uyarı", MessageBoxButtons.YesNo);
            if (cikis == DialogResult.Yes)
            {
                con.Open();
                string bag = "DELETE FROM Oykod WHERE İsim_table=@İsim_table";
                OleDbCommand komut = new OleDbCommand(bag, con);
                komut.Parameters.AddWithValue("@İsim_table", textBox7.Text);
                komut.ExecuteNonQuery();
                con.Close();

                for (int i = 0; i < dataGridView4.Rows.Count; i++)
                {
                    textBox7.Text = dataGridView4.Rows[i].Cells[0].Value.ToString();

                    con.Open();
                    string bag2 = "DELETE FROM kontrol WHERE sonuclar=@sonuclar";
                    OleDbCommand komut2 = new OleDbCommand(bag2, con);
                    komut2.Parameters.AddWithValue("@sonuclar", textBox7.Text);
                    komut2.ExecuteNonQuery();
                    con.Close();
                }
                button6.PerformClick();
                griddoldur();
                label13.Text = dataGridView2.Rows.Count.ToString();

                string x = "Birden fazla aynı e-mail kullanımı tespit edilmiştir: " + textBox12.Text;
                //Burada iki durum var aslında kandırmak için
                //1. Kişi önce bi kendi normal oyu için doldurdu bitti daha sonra aynı e-mail ve farklı bir isimle kod istiyorsa eğer sanırsam bunun hileden farklı bir açıklaması olmayacaktır AMA, ama diyorum çünkü isim olarak birbirine benzeyen kişilerin hacettepeli e-mailleri de haliyle birbirine benzeyecektir ve bu durumdan dolayı bir karışma ihtimali olabilir, bunu da göz önünde bulundurmak lazım; yok eğer isimler ve uzantılı e-mail alakasız ise mutlaka orada bir hinlik vardır
                //2. ve belki de daha masum gibi görünen bir aldatma türü var ve çok pis çamura yatabilir; farklı gmail hesapları ile giriş yapıp aynı bilgileri doldurarak aa pardon ya ben iki kere doldurmuşum kusura bakmayın diyebilir ve oyunun geçerli sayılabilmesi durumu olabilir
                //ama tabi yine de bir sorun varsa kişinin kendisi ile görüşmek gerekir çünkü arkaplanını bilmeden birinin oyunu geçersiz saymak çok salakça olur

                con.Open();
                string bag3 = "insert into gecersiz(isim,sebep) values (@isim,@sebep)";
                OleDbCommand komut3 = new OleDbCommand(bag3, con);
                komut3.Parameters.AddWithValue("@isim", textBox7.Text);
                komut3.Parameters.AddWithValue("@sebep", x);
                komut3.ExecuteNonQuery();
                con.Close();
                griddoldur3();

                int y = Convert.ToInt32(label29.Text.ToString());
                y++;
                string b = Convert.ToString(y);
                label29.Text = b;
            }
            if (cikis == DialogResult.No)
            {
                
            }
        }
        private void dataGridView3_MouseClick(object sender, MouseEventArgs e)
        {
            textBox7.Text = dataGridView3.CurrentRow.Cells[0].Value.ToString();
        }
        private void dataGridView3_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                DialogResult cikis = new DialogResult();
                cikis = MessageBox.Show("Tıkladığınız kişiyi silmek istediğinize emin misiniz? ?", "Uyarı", MessageBoxButtons.YesNo);
                if (cikis == DialogResult.Yes)
                {
                    int index = dataGridView3.CurrentRow.Index;
                    dataGridView3.Rows.RemoveAt(index);

                    textBox4.Text = dataGridView3.CurrentRow.Cells[0].Value.ToString();
                    DataView dv = ds2.Tables[0].DefaultView;
                    dv.RowFilter = "Kod_table Like '%" + textBox4.Text + "%' or Email_table Like '%" + textBox4.Text + "%'";
                    dataGridView2.DataSource = dv;
                    griddoldur();

                    if (dataGridView2.Rows.Count == 0)
                    {
                        txt_sebepbelirt.Text = "Bulunamadı";
                    }
                    else
                    {
                        txt_sebepbelirt.Text = dataGridView2.Rows[0].Cells[2].ToString();
                    }

                    con.Open();
                    string bag = "DELETE FROM Oykod WHERE İsim_table=@İsim_table";
                    OleDbCommand komut = new OleDbCommand(bag, con);
                    komut.Parameters.AddWithValue("@İsim_table", textBox7.Text);
                    komut.ExecuteNonQuery();
                    con.Close();

                    string x = "Sistemin gönderdiği kodlardan farklı olan bir kod girişi yaptığınız için oyunuz geçersiz kabul edilmiştir. Size gönderilen doğrulama kodu: " + txt_sebepbelirt.Text + " Sizin girdiğiniz doğrulama kodu: " + textBox4.Text;

                    con.Open();
                    string bag2= "insert into gecersiz(isim,sebep) values (@isim,@sebep)";
                    OleDbCommand komut2 = new OleDbCommand(bag2, con);
                    komut2.Parameters.AddWithValue("@isim", textBox7.Text);
                    komut2.Parameters.AddWithValue("@sebep",x);
                    komut2.ExecuteNonQuery();
                    con.Close();
                    griddoldur3();

                    listBox6.Items.Clear();
                    button2.PerformClick();
                    label16.Text = dataGridView3.Rows.Count.ToString();

                    int y = Convert.ToInt32(label29.Text.ToString());
                    y++;
                    string b = Convert.ToString(y);
                    label29.Text = b;
                }
                if (cikis == DialogResult.No)
                {

                }
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            string x = "Oy kullanmak için geçersiz olan bir email uzantısı kullanıldığı tespit edilmiştir: " + textBox9.Text;
            con.Open();
            string bag3 = "insert into gecersiz(isim,sebep) values (@isim,@sebep)";
            OleDbCommand komut3 = new OleDbCommand(bag3, con);
            komut3.Parameters.AddWithValue("@isim", textBox10.Text);
            komut3.Parameters.AddWithValue("@sebep", x);
            komut3.ExecuteNonQuery();
            con.Close();
            griddoldur3();

            int y = Convert.ToInt32(textBox11.Text);
            dataGridView1.Rows.RemoveAt(y);

            listBox2.Items.Clear();
            listBox5.Items.Clear();
            listBox7.Items.Clear();
            button8.PerformClick();
            label10.Text = dataGridView1.Rows.Count.ToString();

            int z = Convert.ToInt32(label29.Text.ToString());
            z++;
            string b = Convert.ToString(z);
            label29.Text = b;
        }
        private void button8_Click(object sender, EventArgs e)
        {
            List<string> isim = new List<string>();
            List<string> mail = new List<string>();
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                //demek önemli olan exceldeki adı neyse onu yazıyormuşsun, veritabının o burada sonuç olarak
                isim.Add(dataGridView1.Rows[i].Cells["isim"].Value.ToString());
                mail.Add(dataGridView1.Rows[i].Cells["email"].Value.ToString());
            }

            string[] isim_ = isim.ToArray();
            string[] mail_ = mail.ToArray();

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if(mail_[i].Contains(textBox8.Text))
                {
                    int index = dataGridView1.Rows[i].Index;
                    listBox2.Items.Add(mail_[i].ToString());
                    listBox5.Items.Add(isim_[i].ToString());
                    listBox7.Items.Add(index);
                }
            }
            if (listBox2.Items.Count == 0)
            {
                listBox2.Items.Add(textBox8.Text + " formatında bir e-mail adresi bulunamamıştır.");
            }
        }

        private void listBox2_Click(object sender, EventArgs e)
        {
            int ind = listBox2.SelectedIndex;
            textBox9.Text = listBox2.Items[ind].ToString();
            textBox10.Text = listBox5.Items[ind].ToString();
            textBox11.Text = listBox7.Items[ind].ToString();
        }

        private void listBox6_Click(object sender, EventArgs e)
        {
            int ind = listBox6.SelectedIndex;
            textBox4.Text = listBox6.Items[ind].ToString();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Form2 x = new Form2();
            x.Show();
        }

        private void dataGridView3_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if ((e.ColumnIndex == 0 || e.ColumnIndex == 2 || e.ColumnIndex == 3) && e.Value != null)
            {
                e.Value = new string('*', e.Value.ToString().Length);
            }
        }
    }
}
