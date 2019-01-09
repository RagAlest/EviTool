using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Coalescence
{
    public partial class Form1 : Form
    {
        private dynamic app;
        private dynamic books;
        private dynamic book;
        private dynamic sheets;
        private dynamic sheet;
        private dynamic range;
        private dynamic shape;

        private static char[] alphabet = Enumerable.Range('A', 'Z' - 'A' + 1).Select(i => (char)i).ToArray();

        public Form1()
        {
            InitializeComponent();
            this.txtTargetExcel.Text = "C:\\Users\\jetbl\\Desktop\\a.xlsx";
            this.txtTargetFolder.Text = "C:\\png";
            this.lblStatus.Text = "Stand-by";
            this.lblStatus.ForeColor = Color.Green;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // 感度良好524
            this.lblStatus.Text = "Execute " + this.txtTargetFolder.Text;
            this.lblStatus.ForeColor = Color.Salmon;
            this.CreateExcelFile(this.txtTargetExcel.Text, this.CreateFileList(this.txtTargetFolder.Text));
        }

        private void CreateExcelFile(string excelFile, string[,] imgList)
        {
            string pic = string.Empty;
            string position = string.Empty;
            try
            {
                Type excelApp = Type.GetTypeFromProgID("Excel.Application");
                app = Activator.CreateInstance(excelApp);

                app.DisplayAlerts = false;

                books = app.WorkBooks;
                book = books.Add();
                sheets = book.Sheets;
                sheet = sheets["Sheet1"];

                for (int y = 0; y < imgList.GetLength(1); ++y)
                {
                    for (int x = 0; x < imgList.GetLength(0); ++x)
                    {
                        position = string.Empty;
                        pic = System.IO.Path.Combine(this.txtTargetFolder.Text, imgList[x, y]);
                        this.txtLog.Text += "pasting " + pic;

                        if(x > 1)
                        {
                            // 1枚の幅 20 * 何列目か / アルファベットの数 = 何週目のアルファベットか
                            // (計算結果(商)が0 = アルファベット1文字, 計算結果(商)が1 = AA-AZ, 2 = BA-BZ)
                            position = alphabet[20 * x / 25 - 1].ToString();
                            position += alphabet[20 * x % 25].ToString();
                        }
                        else
                        {
                            position += alphabet[20 * x % 25].ToString();
                        }

                        this.txtLog.Text += " to Cell" + string.Concat(position, y * 15) + "..." + Environment.NewLine;

                        range = sheet.Range[string.Concat(position, (y * 25 + 1).ToString())];

                        double left = range.Left;
                        double top = range.Top;
                        double width = 0;
                        double height = 0;

                        shape = sheet.Shapes.AddPicture(pic, true, true, left, top, width, height);

                        shape.ScaleHeight(1.0, true);
                        shape.ScaleWidth(1.0, true);
                    }
                }

                book.SaveAs(excelFile);
                book.Close();
                app.Quit();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.StackTrace);
            }
            finally
            {
                this.releaseObject((object)shape);
                this.releaseObject((object)range);
                this.releaseObject((object)sheet);
                this.releaseObject((object)sheets);
                this.releaseObject((object)book);
                this.releaseObject((object)books);
                this.releaseObject((object)app);
                this.lblStatus.Text = "Stand-by";
                this.lblStatus.ForeColor = Color.Green;
            }
        }

        private string[,] CreateFileList(string folderName)
        {
            System.IO.DirectoryInfo directory = new System.IO.DirectoryInfo(folderName);
            System.IO.FileInfo[] pngFiles = directory.GetFiles("*.png", System.IO.SearchOption.AllDirectories);
            
            //どっちか、もしくはいい感じのやつ
            //var orderedFiles = pngFiles.OrderBy(x => x.FullName);
            var orderedFiles = pngFiles.OrderBy(x => x.CreationTime);

            int columnCount = Convert.ToInt32(this.txtColumns.Text);
            //int rowCount = Convert.ToInt32(pngFiles.Length / columnCount);
            int rowCount = Convert.ToInt32(orderedFiles.Count() / columnCount);
            int fileCount = 0;

            string[,] imgList = new string[columnCount, rowCount];

            for (int y = 0; y < rowCount; ++y)
            {
                for (int x = 0; x < columnCount; ++x)
                {
                    imgList[x, y] = orderedFiles.ToArray()[fileCount].Name;//
                    fileCount++;
                }
            }

            return imgList;
        }

        private void releaseObject(object obj)
        {
            try
            {
                Marshal.FinalReleaseComObject(obj);
                obj = null;
            }
            catch (Exception e)
            {
                obj = null;
                MessageBox.Show("COMオブジェクト解放失敗 : " + e.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
