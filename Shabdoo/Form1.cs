using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace Shabdoo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
           
        }
        // Определение переменной oWord
        Word._Application oWord = new Word.Application();
        void button1_Click(object sender, EventArgs e)
        {
            //var application = new Microsoft.Office.Interop.Word.Application();
            //var document = new Microsoft.Office.Interop.Word.Document();
            _Document oDoc = GetDoc(@"D:\Проект работы с шаблонами\NEW_C_PRJ\test.docx");
            oDoc.SaveAs2(FileName:@"D:\Проект работы с шаблонами\NEW_C_PRJ\test_2.docx");
            oDoc.Close();
        }
        private _Document GetDoc(string path)
        {
            _Document oDoc = oWord.Documents.Add(path);
            SetTemplate(oDoc);
            return oDoc;
        }
        // Замена закладки SECONDNAME на данные введенные в textBox
        private void SetTemplate(Word._Document oDoc)
        {
            oDoc.Bookmarks["SECONDNAME"].Range.Text = "ТЕСТ";
            // если нужно заменять другие закладки, тогда копируем верхнюю строку изменяя на нужные параметры 

        }
    }
}
