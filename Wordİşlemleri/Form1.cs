using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using wordektar = Microsoft.Office.Interop.Word; //Word kütüphanemiz, referans verdikten sonra yazılmalı.
using System.Reflection; //Tüm sınıfları,dataları vs bilgi olarak kendisinde tutar.



namespace Wordİşlemleri
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            wordektar.Application wordapp = new wordektar.Application();
            wordapp.Visible = true;
            wordektar.Document worddoc;
            object wordobj = System.Reflection.Missing.Value;
            worddoc = wordapp.Documents.Add(ref wordobj);
            wordapp.Selection.TypeText(richTextBox1.Text);
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form2 frm = new Form2();
            frm.Show();
            this.Hide();
        }
    }
}
