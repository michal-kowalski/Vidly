using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Threading;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SymWord
{
	public partial class Form1 : Form
	{
		

		private void Form1_Load(object sender, EventArgs e)
		{

		}

		private void button1_Click(object sender, EventArgs e) {
			Open("E:\\CorrespondenceSticker.docx");
		}

		public object missing = System.Reflection.Missing.Value;
		private object filePath;
		private bool hasChanged;

		public int ProcessID = -1;

		private bool freezeEvents = false;

		public Dictionary<string, object> UserData;

		public bool HasChanged
		{
			get { return hasChanged; }
			set { hasChanged = value; }
		}


		public object FilePath
		{
			get { return filePath; }
			//set { filePath = value; }
		}

		private Microsoft.Office.Interop.Word.Document document = null;
		private Microsoft.Office.Interop.Word.Application app = null;

		public event System.EventHandler OnWordExit;
		public event System.EventHandler<DocBeforeSaveEventArgs> OnWordSave;
		public event System.EventHandler<DocBeforeCloseEventArgs> OnDocClose;


		public Microsoft.Office.Interop.Word.Application WordAplication
		{
			get { return app; }
		}

		public Microsoft.Office.Interop.Word.Document Document
		{
			get { return document; }
		}

		public Form1()
		{
			InitializeComponent();

			DateTime startTime = DateTime.Now;

			if (app == null)
			{

				app = new Microsoft.Office.Interop.Word.Application();

				//Microsoft.Office.Core.COMAddIns comAddins = app.COMAddIns;


				//foreach (Microsoft.Office.Core.COMAddIn item in comAddins)
				//{
				//    System.Windows.Forms.MessageBox.Show(item.Description);
				//    item.Connect = true;
				//}





				//for (int i = 0; i < app.COMAddIns.Count; i++)
				//{

				//}



			}

			DateTime endTime = DateTime.Now;

			foreach (System.Diagnostics.Process pr in System.Diagnostics.Process.GetProcessesByName("WINWORD"))
			{
				try
				{
					if ((pr.StartTime >= startTime) && (pr.StartTime <= endTime))
					{
						ProcessID = pr.Id;
						break;
					}
				}
				catch// (Exception ex)
				{
				}
			}

			this.UserData = new Dictionary<string, object>();


			//app.ApplicationEvents2_Event_Quit += new Microsoft.Office.Interop.Word.ApplicationEvents2_QuitEventHandler(app_ApplicationEvents2_Event_Quit);
			//app.ApplicationEvents2_Event_DocumentBeforeSave += new Microsoft.Office.Interop.Word.ApplicationEvents2_DocumentBeforeSaveEventHandler(app_ApplicationEvents2_Event_DocumentBeforeSave);
			//app.ApplicationEvents2_Event_Docu

			app.DocumentBeforeSave += new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(app_ApplicationEvents2_Event_DocumentBeforeSave);
			app.DocumentBeforeClose += new Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(app_ApplicationEvents2_Event_DocumentBeforeClose);
			hasChanged = false;



		}



		void app_ApplicationEvents2_Event_DocumentBeforeClose(Microsoft.Office.Interop.Word.Document Doc, ref bool Cancel)
		{
			try
			{
				if (!freezeEvents)
					if (Doc.Name.Equals(document.Name))
						if (OnDocClose != null)
						{
							freezeEvents = true;
							DocBeforeCloseEventArgs args = new DocBeforeCloseEventArgs();
							OnDocClose(this, args);
							Cancel = args.Cancel;
							freezeEvents = false;
						}
			}
			catch { }
		}

		void app_ApplicationEvents2_Event_DocumentBeforeSave(Microsoft.Office.Interop.Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
		{
			try
			{
				if (!freezeEvents)
					if (Doc.Name.Equals(document.Name))
						if (OnWordSave != null)
						{
							DocBeforeSaveEventArgs args = new DocBeforeSaveEventArgs();
							args.SaveAs = SaveAsUI;
							OnWordSave(this, args);
							Cancel = args.Cancel;
						}
			}
			catch { }
		}

		void app_ApplicationEvents2_Event_Quit()
		{
			if (!freezeEvents)
				if (OnWordExit != null) OnWordExit(this, new EventArgs());
		}

		public void CreateDoc(string fileName)
		{
			this.filePath = fileName;

			object newTemplate = false;
			object docType = 0;
			object isVisible = true;
			object optional = Missing.Value;

			document = app.Documents.Add(ref this.filePath, ref optional, ref optional, ref isVisible);
			//document = app.Documents.AddOld(ref this.filePath, ref newTemplate);
			hasChanged = false;


		}


		public string CheckSpelling(string text, Microsoft.Office.Interop.Word.WdLanguageID lang)
		{


			this.document.Words.First.InsertBefore(text);

			object first = 0;
			object last = this.document.Characters.Count - 1;

			this.document.Range(ref first, ref last).LanguageID = lang;

			Word.ProofreadingErrors spellErrorsColl = this.document.SpellingErrors;

			object optional = Missing.Value;

			//app.ShowMe();

			//ShowRightWindow srw = new ShowRightWindow();

			//System.Windows.Forms.MessageBox.Show(WordAplication.Version.ToString());


			//switch (lang)
			//{
			//	case Microsoft.Office.Interop.Word.WdLanguageID.wdPolish:
			//		srw.WindowsCaptions.Add("Pisownia: Polski");
			//		break;
			//	case Microsoft.Office.Interop.Word.WdLanguageID.wdEnglishUK:
			//		srw.WindowsCaptions.Add("Pisownia: Angielski (Zjednoczone Królestwo)");
			//		srw.WindowsCaptions.Add("Pisownia: Angielski (Wielka  Brytania)");
			//		srw.WindowsCaptions.Add("Pisownia: Angielski (Wielka Brytania)");
			//		break;
			//}

			//Thread th = new Thread(new ThreadStart(srw.AltTab));
			//th.Start();

			this.document.CheckSpelling(
				ref optional, ref optional, ref optional, ref optional, ref optional, ref optional,
				ref optional, ref optional, ref optional, ref optional, ref optional, ref optional);


			last = this.document.Characters.Count - 1;

			string temp = this.document.Range(ref first, ref last).Text;

			if (temp == null) temp = "";

			this.document.Range(ref first, ref last).Delete(ref optional, ref optional);

			temp = temp.Replace("\r", System.Environment.NewLine);

			return temp;
		}

		public void SendAsMail()
		{
			//document.SendMail();

			//object subject = "Monit";
			//object show = true;
			//document.SendForReview(ref missing, ref subject, ref show, ref missing);             
			this.SendAsMail("");

		}

		public void SendAsMail(string to)
		{
			//document.SendMail();

			object subject = "Monit";
			object recipients = to;
			object show = true;
			document.SendForReview(ref recipients, ref subject, ref show, ref missing);

		}

		public void CreateNew(string fileName)
		{
			this.filePath = fileName;

			object newTemplate = false;
			object docType = 0;
			object isVisible = false;


			document = app.Documents.Add(ref missing, ref newTemplate, ref docType, ref isVisible);
			this.SaveAs(fileName);

			hasChanged = false;

		}

		public void CreateNew()
		{
			this.filePath = "";
			object template = Missing.Value;
			object newTemplate = Missing.Value;
			object documentType = Missing.Value;
			object isVisible = false;

			document = app.Documents.Add(ref template, ref newTemplate, ref documentType, ref isVisible);
		}


		public void Open(string fileName)
		{
			this.Open(fileName, false);
		}

		public void Open(string fileName, bool readOnly)
		{
			this.filePath = fileName;
			object fn = (object)fileName;
			object oreadOnly = readOnly;
			object objtrue = true;
			object objfalse = false;
			object format = Microsoft.Office.Interop.Word.WdOpenFormat.wdOpenFormatAllWord;

			//document = app.Documents.Open(ref fn, ref missing, ref oreadOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
			if (app.Version.StartsWith("12.") || app.Version.StartsWith("14."))
				//document = app.Documents.Open(ref fn, ref missing, ref oreadOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref objtrue, ref missing, ref missing);
				document = app.Documents.Open(ref fn, ref missing, ref oreadOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
			else
				document = app.Documents.Open(ref fn, ref objfalse, ref oreadOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
			hasChanged = false;
		}

		public void PrintToFile(string fileName, string printerName)
		{
			object copies = 1;
			object ofile = fileName;
			object otrue = true;
			string printer = app.ActivePrinter;

			app.ActivePrinter = printerName;
			document.PrintOut(ref missing, ref missing, ref missing, ref ofile, ref missing, ref missing, ref missing, ref copies, ref missing, ref missing, ref otrue, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
			app.ActivePrinter = printer;

		}

		public void Show()
		{
			if (document == null) return;
			app.Visible = true;
			document.Activate();
			app.ShowMe();
		}


		public void Hide()
		{
			app.Visible = false;
		}

		public void Close()
		{
			object saveCh = (object)false;
			this.CloseDoc();
			app.Quit(ref saveCh, ref missing, ref missing);
			app = null;
			try
			{
				if (ProcessID > 0) System.Diagnostics.Process.GetProcessById(ProcessID).Kill();
			}
			catch { }

		}

		public void CloseWithoutSave()
		{
			object saveCh = (object)false;
			this.CloseDocWithoutSave();
			app.Quit(ref saveCh, ref missing, ref missing);
			app = null;

			try
			{
				if (ProcessID > 0) System.Diagnostics.Process.GetProcessById(ProcessID).Kill();
			}
			catch { }

		}

		public void CloseDocWithoutSave()
		{
			object saveCh = (object)false;
			document.Close(ref saveCh, ref missing, ref missing);
		}


		public void CloseDoc()
		{
			object saveCh = (object)true;
			filePath = document.Path + @"\" + document.Name;
			//document.SaveAs(ref filePath);
			document.Close(ref saveCh, ref missing, ref missing);
			//document.Application.Documents.Close(ref saveCh, ref missing, ref missing);
		}

		public void Replace(string searchFor, string replaceWith)
		{
			if (document == null) throw new Exception("Create document first");

			foreach (Microsoft.Office.Interop.Word.Paragraph para in document.Paragraphs)
			{
				if (para.Range.Text.Contains(searchFor))
				{
					para.Range.Text = para.Range.Text.Replace(searchFor, replaceWith);
				}

			}


		}

		public void SaveAs(string fileName)
		{
			object fn = (object)fileName;
			Document.SaveAs(ref fn, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
		}

		public void Save()
		{
			object fn = (object)this.FilePath;
			Document.SaveAs(ref fn);
			filePath = document.Path + @"\" + document.Name;
		}

		public void Save(bool withEvent)
		{
			bool oldState = freezeEvents;
			freezeEvents = !withEvent;
			this.Save();
			freezeEvents = oldState;
		}


		public bool SetBookmark(string name, string value)
		{
			if (document == null) throw new Exception("Create document first");

			object oBookMark = name;
			if (document.Bookmarks.Exists(name))
			{
				document.Bookmarks.get_Item(ref oBookMark).Range.Text = value;
				return true;
			}
			else
			{
				return false;
			}

		}

		public bool SetBookmarkMulti(string name, string value)
		{
			if (document == null) throw new Exception("Create document first");

			SetBookmark(name, value);

			for (int i = 0; i < 10; i++)

				SetBookmark(name + i.ToString(), value);

			return true;

		}

		public void ShowOnTop()
		{
			app.ShowMe();
		}

		public class DocBeforeCloseEventArgs : EventArgs
		{
			public bool Cancel = false;
		}

		public class DocBeforeSaveEventArgs : EventArgs
		{
			public bool Cancel = false;
			public bool SaveAs = false;
		}

		internal void SaveAsPDF(string pdfFile)
		{
			Document.ExportAsFixedFormat(pdfFile, Word.WdExportFormat.wdExportFormatPDF);
		}

	}

	
}
