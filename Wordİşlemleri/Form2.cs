using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using wordaktar = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace Wordİşlemleri
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            object omissing = System.Reflection.Missing.Value; 
            object dokumansonu = "\\endofdoc";
            wordaktar.Application olustur;
            wordaktar.Document icerik;

            olustur = new wordaktar.Application();
            olustur.Visible = true;
            icerik = olustur.Documents.Add(ref omissing);
            //olustur.Selection.TypeText(richTextBox1.Text);


            wordaktar.Paragraph paragraf1;
            paragraf1 = icerik.Content.Paragraphs.Add(ref omissing);
            paragraf1.Range.Text = "C# derslerinde bugün Word işlemleri yaptık.";
            paragraf1.Range.Font.Bold = 1;
            paragraf1.Format.SpaceAfter = 20;
            paragraf1.Range.InsertParagraphAfter();


            wordaktar.Paragraph paragraf2;
            object orng = icerik.Bookmarks.get_Item(ref dokumansonu).Range; //üstteki paragrafın devamı olarak.
            paragraf2 = icerik.Content.Paragraphs.Add(ref omissing);
            paragraf2.Range.Text = "Yarın da Excel çalışacağız.";
            paragraf2.Range.Font.Italic = 1;
            paragraf2.Range.Font.Bold = 5;
            paragraf2.Format.SpaceAfter = 50;
            paragraf2.Range.InsertParagraphAfter();

            /*wordaktar.Paragraph paragraf3;
            orng = icerik.Bookmarks.get_Item(ref dokumansonu).Range;
            paragraf3 = icerik.Content.Paragraphs.Add(ref omissing);
            paragraf3.Range.Text = richTextBox1.Text;
            paragraf3.Range.Font.Bold = 10;
            paragraf3.Format.SpaceAfter = 100;
            paragraf3.Range.Underline = 0;
            paragraf3.Range.InsertParagraphAfter();*/
            
            //matris oluşturma.
            wordaktar.Table olusturtablo;
            wordaktar.Range wrdng = icerik.Bookmarks.get_Item(ref dokumansonu).Range;
            olusturtablo = icerik.Tables.Add(wrdng, 3, 5, ref omissing, ref omissing); //3 e 5'lik bir matris oluşturdu.
            olusturtablo.Range.ParagraphFormat.SpaceAfter = 10;

            int r, c;
            string strText;

            for(r=1;r<=3;r++)
                for(c=1;c<=5;c++)
                {
                    strText = "Satır" + r + "Sütun" + c;
                    olusturtablo.Cell(r, c).Range.Text = strText;
                    olusturtablo.Rows[1].Range.Font.Bold = 1;
                    olusturtablo.Rows[1].Range.Font.Italic = 1;
                    

                }

           

        }
    }
}
