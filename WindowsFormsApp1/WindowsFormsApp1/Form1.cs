using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WindowsFormsApp1
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void Form1_Load(object sender, EventArgs e)
		{

		}
		private void button1_Click_1(object sender, EventArgs e)

		{

			object oMissing = System.Reflection.Missing.Value;

			object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

			object filename = @"E:\\CorrespondenceSticker.docx";

			//Start Word and create a new document.

			Word._Application oWord;

			Word._Document oDoc;

			oWord = new Word.Application();

			oWord.Visible = true;

			oDoc = oWord.Documents.Add(ref filename, ref oMissing,

			ref oMissing, ref oMissing);

			//doc.Open()= “E:/lakshmanan.doc”;

			Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

			//Insert a paragraph at the beginning of the document.

			Word.Paragraph oPara1;

			oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);

			oPara1.Range.Text = "MANIVANNAN";

			oPara1.Range.Font.Bold = 1;

			oPara1.Range.Font.Size = 14;

			//24 pt spacing after paragraph.
		}

		
		
	}
}
