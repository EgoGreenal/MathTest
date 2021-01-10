using System;
using System.Windows.Forms;
using MSWord = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace MathTest
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}
		private void button1_Click(object sender, EventArgs e)
		{
            progressBar1.Value = 0;
            int n, m;
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
            if (n < 2 || m < 1)
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
			//wordApp.Visible = true;
            object unite = MSWord.WdUnits.wdStory;
            wordDoc.PageSetup.PaperSize = MSWord.WdPaperSize.wdPaperA4;//设置纸张样式为A4纸
            wordDoc.PageSetup.Orientation = MSWord.WdOrientation.wdOrientPortrait;//排列方式为垂直方向
            wordDoc.PageSetup.TopMargin = 57.0f;
            wordDoc.PageSetup.BottomMargin = 57.0f;
            wordDoc.PageSetup.LeftMargin = 57.0f;
            wordDoc.PageSetup.RightMargin = 57.0f;
            #region 添加表格、填充数据、设置表格行列宽高、合并单元格、添加表头斜线、给单元格添加图片
            //wordDoc.Content.InsertAfter("\n");//这一句与下一句的顺序不能颠倒，原因还没搞透
            wordApp.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphLeft;
            //object WdLine2 = MSWord.WdUnits.wdLine;//换一行;  
            //wordApp.Selection.MoveDown(ref WdLine2, 6, ref Nothing);//向下跨15行输入表格，这样表格就在文字下方了，不过这是非主流的方法

            wordDoc.Paragraphs.Last.Range.Font.Name = "Proxy 3";
            wordDoc.Paragraphs.Last.Range.Font.Bold = 1;
            //设置表格的行数和列数
            int tableRow = 20;
            int tableColumn = 5 * 2 + 1;

            for (int cas = 0; cas < m; ++cas)
            {
                if (cas > 0) { wordApp.Selection.EndKey(ref unite, ref Nothing); wordDoc.Content.InsertAfter("\n"); }
                wordApp.Selection.EndKey(ref unite, ref Nothing); //将光标移动到文档末尾
                //定义一个Word中的表格对象
                MSWord.Table table = wordDoc.Tables.Add(wordApp.Selection.Range, tableRow, tableColumn, ref Nothing, ref Nothing);

                //默认创建的表格没有边框，这里修改其属性，使得创建的表格带有边框 
                //表格的索引是从1开始的。
                //wordDoc.Tables[1].Cell(1, 1).Range.Text = "列\n行";
                //设置table样式
                table.Borders.Enable = 1;//这个值可以设置得很大，例如5、13等等
                table.Rows.HeightRule = MSWord.WdRowHeightRule.wdRowHeightAtLeast;//高度规则是：行高有最低值下限？
                table.Rows.Height = wordApp.CentimetersToPoints(0.8F);// 

                table.Range.Font.Name = "Proxy 3";
                table.Range.Font.Size = 20F;
                table.Range.Font.Bold = 1;

                table.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;//表格文本居中
                table.Range.Cells.VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;//文本垂直贴到底部
                //设置table边框样式
                //table.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleDouble;//表格外框是双线
                //table.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;//表格内框是单线
                table.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleNone;
                table.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleNone;

                //table.Rows[1].Range.Font.Bold = 1;//加粗
                //table.Rows[1].Range.Font.Size = 12F;
                //table.Cell(1, 1).Range.Font.Size = 10.5F;
                wordApp.Selection.Cells.Height = 35;//所有单元格的高度

                Random rnd = new Random();
                for (int i = 1; i <= tableRow; i++)
                {
                    for (int j = 1, a, b, c; j <= 2; j++)
                    {
                        do
                        {
                            a = rnd.Next(0, n);
                            b = rnd.Next(0, n);
                        } while (a + b > n);
                        c = rnd.Next(1, 8);
                        switch (c)
                        {
                            case 1:
                                table.Cell(i, j * 6 - 5).Range.Text = a.ToString();
                                table.Cell(i, j * 6 - 4).Range.Text = "+";
                                table.Cell(i, j * 6 - 3).Range.Text = b.ToString();
                                table.Cell(i, j * 6 - 2).Range.Text = "=";
                                table.Cell(i, j * 6 - 1).Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                                break;
                            case 2:
                                table.Cell(i, j * 6 - 5).Range.Text = a.ToString();
                                table.Cell(i, j * 6 - 4).Range.Text = "+";
                                table.Cell(i, j * 6 - 3).Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                                table.Cell(i, j * 6 - 2).Range.Text = "=";
                                table.Cell(i, j * 6 - 1).Range.Text = (a + b).ToString();
                                break;
                            case 3:
                                table.Cell(i, j * 6 - 5).Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                                table.Cell(i, j * 6 - 4).Range.Text = "+";
                                table.Cell(i, j * 6 - 3).Range.Text = b.ToString();
                                table.Cell(i, j * 6 - 2).Range.Text = "=";
                                table.Cell(i, j * 6 - 1).Range.Text = (a + b).ToString();
                                break;
                            case 4:
                                table.Cell(i, j * 6 - 5).Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                                table.Cell(i, j * 6 - 4).Range.Text = "-";
                                table.Cell(i, j * 6 - 3).Range.Text = a.ToString();
                                table.Cell(i, j * 6 - 2).Range.Text = "=";
                                table.Cell(i, j * 6 - 1).Range.Text = b.ToString();
                                break;
                            case 5:
                                table.Cell(i, j * 6 - 5).Range.Text = (a + b).ToString();
                                table.Cell(i, j * 6 - 4).Range.Text = "-";
                                table.Cell(i, j * 6 - 3).Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                                table.Cell(i, j * 6 - 2).Range.Text = "=";
                                table.Cell(i, j * 6 - 1).Range.Text = b.ToString();
                                break;
                            case 6:
                                table.Cell(i, j * 6 - 5).Range.Text = (a + b).ToString();
                                table.Cell(i, j * 6 - 4).Range.Text = "-";
                                table.Cell(i, j * 6 - 3).Range.Text = b.ToString();
                                table.Cell(i, j * 6 - 2).Range.Text = "=";
                                table.Cell(i, j * 6 - 1).Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                                break;
                            case 7:
                                c = rnd.Next(0, 2);
                                table.Cell(i, j * 6 - 5).Range.Text = (c > 0 ? a + b : a).ToString();
                                table.Cell(i, j * 6 - 4).Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
                                table.Cell(i, j * 6 - 3).Range.Text = b.ToString();
                                table.Cell(i, j * 6 - 2).Range.Text = "=";
                                table.Cell(i, j * 6 - 1).Range.Text = (c == 0 ? a + b : a).ToString();
                                break;
                        }
                    }
                    progressBar1.Value = (100 * (cas * tableRow + i) + (tableRow * m - 1) / 2) / (tableRow * m);
                }
            }

/*
            //添加行
            table.Rows.Add(ref Nothing);
            table.Rows[tableRow + 1].Height = 35;//设置新增加的这行表格的高度
            //向新添加的行的单元格中添加图片
            string FileName = Environment.CurrentDirectory + "\\6.jpg";//图片所在路径
            object LinkToFile = false;
            object SaveWithDocument = true;
            object Anchor = table.Cell(tableRow + 1, tableColumn).Range;//选中要添加图片的单元格
            wordDoc.Application.ActiveDocument.InlineShapes.AddPicture(FileName, ref LinkToFile, ref SaveWithDocument, ref Anchor);

            wordDoc.Application.ActiveDocument.InlineShapes[1].Width = 50;//图片宽度
            wordDoc.Application.ActiveDocument.InlineShapes[1].Height = 35;//图片高度

            // 将图片设置为四周环绕型
            MSWord.Shape s = wordDoc.Application.ActiveDocument.InlineShapes[1].ConvertToShape();
            s.WrapFormat.Type = MSWord.WdWrapType.wdWrapSquare;
*/

            //除第一行外，其他行的行高都设置为20
            //for (int i = 2; i <= tableRow; i++)
            //{
            //    table.Rows[i].Height = 20;
            //}

            //将表格左上角的单元格里的文字（“行” 和 “列”）居右
            //table.Cell(1, 1).Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphRight;
            //将表格左上角的单元格里面下面的“列”字移到左边，相比上一行就是将ParagraphFormat改成了Paragraphs[2].Format
            //table.Cell(1, 1).Range.Paragraphs[2].Format.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphLeft;

            //table.Columns[1].Width = 50;//将第 1列宽度设置为50

            //将其他列的宽度都设置为75
           // for (int i = 2; i <= tableColumn; i++)
            //{
            //    table.Columns[i].Width = 75;
            //}


            //添加表头斜线,并设置表头的样式
            //table.Cell(1, 1).Borders[MSWord.WdBorderType.wdBorderDiagonalDown].Visible = true;
            //table.Cell(1, 1).Borders[MSWord.WdBorderType.wdBorderDiagonalDown].Color = MSWord.WdColor.wdColorRed;
            //table.Cell(1, 1).Borders[MSWord.WdBorderType.wdBorderDiagonalDown].LineWidth = MSWord.WdLineWidth.wdLineWidth150pt;

            //合并单元格
            //table.Cell(4, 4).Merge(table.Cell(4, 5));//横向合并

            //table.Cell(2, 3).Merge(table.Cell(4, 3));//纵向合并，合并(2,3),(3,3),(4,3)

            #endregion
            //WdSaveFormat为Word 2003文档的保存格式
            object format = MSWord.WdSaveFormat.wdFormatDocumentDefault;// office 2007就是wdFormatDocumentDefault
            //将wordDoc文档对象的内容保存为DOCX文档
            MessageBox.Show("完成");
            wordDoc.SaveAs(ref path, ref format, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            //关闭wordDoc文档对象
            //看是不是要打印
            //wordDoc.PrintOut();
            wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
            //关闭wordApp组件对象
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
        }
    }
}
