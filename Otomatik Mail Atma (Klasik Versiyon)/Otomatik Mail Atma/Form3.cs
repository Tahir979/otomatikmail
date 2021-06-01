using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Otomatik_Mail_Atma
{
    public partial class Form3 : MetroFramework.Forms.MetroForm
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            //satır sildirme olayına dikkat etmek lazım, yoksa hafızada kalıyor kod

            //her bir mail için yarım saat timeout olmuş oldu böylece, ama try catch ekleyelim yine de bakalım o timeout hatasını ne ile yakalıyorsa artık
            //gönderilen mailler yüzde 99 olarak geri alınamayacaktır bu yüzden işleminizden kesim emin olduktan sonra toplu maili gönderme işleminizi başlatın


            //Yok abi sağlıksız oluyor eğer otomatik olarak program almaya kalkışırsa, hadi iki tane e mail ve isim kısmı oldu filan
            //son seçili kalan tab değerini de silelim herhangi bir tuş üstünde seçili kalmasın
            //abi gerçekten çok saçma; nasıl dizi değer tutamıyor da listbox değer tutabiliyor bu ne sikim iş lan çok aşırı saçma
            //access setupu da kurması gerektiğini söyleyelim programı başka bilgisayara kuracak olanlara, o kurulmadan program çalışmaz çünkü
           
            //neden less secure app özelliği açık olmak zorunda
            //setupta kaynak dosyaları da yükleyim ve metrogrid temasının yüklenmesi gerektiğinden filan bahsedeyim yoksa tasarım kısmının açılmayacağından filan
            //zip dosyaları kabul edilmiyor sanırım herhalde tüm dosya türlerini denemek lazım hangisine geçiş veriyor diye
            //hmm yani excel başlığı da önemli değil yani önemli olan 1.sütunda isim 2.sütunda ise e mail olması
            #region
            /*f++;

            long toplam2 = 0;
            foreach (long sayi in boyut)
            {
                toplam2 += sayi;
            }

            long[] dizi = new long[f];
            dizi[f - 1] = toplam2;//demek burası tutamıyor degeri
            listBox4.Items.Add(dizi[f - 1]);

            long toplam = 0;
            long[] dizi2 = new long[listBox4.Items.Count];
            for (int i = 0; i < listBox4.Items.Count; i++)
            {
                toplam += Convert.ToInt32(listBox4.Items[i].ToString());
                dizi2[i] = Convert.ToInt32(listBox4.Items[i].ToString());
            }

            if (toplam >= sinir)
            {
                listBox4.Items.Clear();
                for (int i = 0; i < dizi2.Length - 1; i++)
                {
                    listBox4.Items.Add(dizi2[i].ToString());
                }

                MessageBox.Show("Toplam dosya boyutu 10 MB'ı aşıyor. Lütfen seçtiğiniz dosyaların toplam 10 MB'ı aşmadığından emin olunuz.", "Dosya Boyutu Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }*/

            #endregion
        }

        private void metroLink1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://steelkiwi.com/blog/30-most-captivating-preloaders-for-website/");
        }

        private void metroLabel3_Click(object sender, EventArgs e)
        {

        }

        private void metroLabel4_Click(object sender, EventArgs e)
        {

        }

        private void metroLabel2_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {

        }

        private void metroLabel2_Click_1(object sender, EventArgs e)
        {

        }
    }
}
