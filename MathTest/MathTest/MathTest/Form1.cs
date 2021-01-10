using System;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace MathTest
{
	public partial class Form1 : Form
	{
		private int n, m;
		private Random rnd = new Random();
		public Form1()
		{
			InitializeComponent();
		}
		public string Generate()
		{
			int lop = rnd.Next(2, n > 9 ? 5 : 7), res = rnd.Next(0, n * n + 1) == 0 ? 0 : rnd.Next(0, n + 1), cnt = 0;
			string[] outp = new string[lop * 2 + 3];
			outp[cnt++] = res.ToString();
			for (int i = 1; i <= lop; ++i)
			{
				int x, op;
				do
				{
					op = 1 - 2 * rnd.Next(0, 2);
					x = rnd.Next(0, n * n + 1) == 0 ? 0 : rnd.Next(0, n + 1);
				} while (res + op * x < rnd.Next(0, 2));
				res += op * x;
				outp[cnt++] = op > 0 ? "+" : "-";
				outp[cnt++] = x.ToString();
			}
			outp[cnt++] = "=";
			outp[cnt++] = res.ToString();
			string fin = "";
			int ra = (rnd.Next(0, 3) > 0 ? lop + 1 : rnd.Next(0, lop + 2)) * 2;
			for (int i = 0; i < cnt; ++i) fin += (i == ra ? (n > 9 ? "___" : "__") : outp[i]) + " ";
			return fin;
		}
		private void button1_Click(object sender, EventArgs e)
		{
            progressBar1.Value = 0;
			try
			{
				n = int.Parse(textBox1.Text);
				m = int.Parse(textBox2.Text);
			}
			catch
			{
				MessageBox.Show("错误！");
				return;
			}
			if (n < 2 || m < 1 || n > 99)
			{
				MessageBox.Show("错误！");
				return;
			}
			sFD.Filter = "DOCX 文档 (*.docx) |*.docx";
			sFD.FilterIndex = 1;
			sFD.RestoreDirectory = true;
			object path;
			if (sFD.ShowDialog() == DialogResult.OK) path = sFD.FileName.ToString(); else return;
			MSWord.Application wordApp = new MSWord.ApplicationClass();
			Object Nothing = Missing.Value;
			MSWord.Document wordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);
			wordApp.Visible = false;
			object unite = MSWord.WdUnits.wdStory;
			wordDoc.PageSetup.PaperSize = MSWord.WdPaperSize.wdPaperA4;
			wordDoc.PageSetup.Orientation = MSWord.WdOrientation.wdOrientPortrait;
			wordDoc.PageSetup.TopMargin = 60.0f;
			wordDoc.PageSetup.BottomMargin = 60.0f;
			wordDoc.PageSetup.LeftMargin = 55.0f;
			wordDoc.PageSetup.RightMargin = 55.0f;
			wordApp.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphLeft;
			wordDoc.Paragraphs.Last.Range.Font.Name = "Fira Code";
			wordDoc.Paragraphs.Last.Range.Font.Size = 25;
			wordDoc.Paragraphs.Last.Range.Font.Bold = 0;
			m *= 23;
			for (int i = 1; i <= m; ++i)
			{
				wordDoc.Content.InsertAfter(Generate());
				if (i < m) wordDoc.Content.InsertAfter("\n");
				progressBar1.Value = (i * 100 + 50) / m;
			}
			object format = MSWord.WdSaveFormat.wdFormatDocumentDefault;
			MessageBox.Show("完成");
			wordDoc.SaveAs(ref path, ref format, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
			wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
			wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
		}
    }
}
