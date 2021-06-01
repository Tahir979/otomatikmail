using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Oy_Sayma_Zoom_
{
    public partial class Oysay : Form
    {
        public Oysay()
        {
            InitializeComponent();
        }

        private void Oysay_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string adet = txt_baskan.Text;
            int sayi = Convert.ToInt32(adet.ToString());

            int i = 0;
            int top = 23;
            int left = 10;

            do
            {
                TextBox txt = new TextBox();
                txt.Location = new Point(left, top * (i + 1));
                groupBox1.Controls.Add(txt);
                i++;
            } while (i < sayi);
        }
    }
}
