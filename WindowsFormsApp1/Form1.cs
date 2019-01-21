using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using NPOI;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using NPOI.XWPF.UserModel;
using NPOI.Util;
using System.Globalization;

namespace excelTool
{
    public partial class Form1 : Form
    {
        public int process { get; private set; }

        public static ThreadStart threadStart;
        public static Thread thread;

        private Form2 myProcessBar = null;//弹出的子窗体(用于显示进度条)
        private delegate void IncreaseHandle(int nValue, bool last);//代理创建
        private IncreaseHandle myIncrease = null;//声明代理，用于后面的实例化代里

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void button1_Changed(object sender, EventArgs e)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.textBox1_Click(sender, e);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.textBox2_Click(sender, e);
        }
        private void button4_Click(object sender, EventArgs e)
        {
            this.textBox3_Click(sender, e);
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();     //显示选择文件对话框
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "excel文件 (*.xls;*.xlsx)|*.xls;*.xlsx";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.textBox1.Text = openFileDialog1.FileName;          //显示文件路径
            }
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.Description = "请选择图片所在文件夹";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (dialog.SelectedPath.Trim() != "")
                    this.textBox2.Text = dialog.SelectedPath.Trim();
            }
        }

        private void textBox3_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.Description = "请选择输出文件夹";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (dialog.SelectedPath.Trim() != "")
                    this.textBox3.Text = dialog.SelectedPath.Trim();
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }


        private void button3_Click(object sender, EventArgs e)
        {
            string fileName = this.textBox1.Text;
            string imagePath = this.textBox2.Text;
            string outPath = this.textBox3.Text;

            if (fileName == "")
            {
                MessageBox.Show("请选择一个Excel文件！");
                return;
            }
            if (imagePath == "")
            {
                MessageBox.Show("请选择图片所在目录！");
                return;
            }
            if (outPath == "")
            {
                MessageBox.Show("请选择输出目录！");
                return;
            }

            //方法一：使用Thread类
            threadStart = new ThreadStart(doIt);//通过ThreadStart委托告诉子线程执行什么方法　
            thread = new Thread(threadStart);
            thread.Start();//启动新线程


        }

        private void doIt()
        {

            string fileName = this.textBox1.Text;
            string imagePath = this.textBox2.Text;
            string outPath = this.textBox3.Text;

            bool last = false;

            IWorkbook workbook = null;  //新建IWorkbook对象  

            FileStream fileStream = null;
            try
            {
                fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            }
            catch (Exception exception)
            {
                MessageBox.Show("出现错误：" + exception.Message);
                return;
            }

            MethodInvoker mi = new MethodInvoker(ShowProcessBar);
            this.BeginInvoke(mi);

            if (fileName.IndexOf(".xlsx") > 0) // 2007版本  
            {
                workbook = new XSSFWorkbook(fileStream);  //xlsx数据读入workbook  
            }
            else if (fileName.IndexOf(".xls") > 0) // 2003版本  
            {
                workbook = new HSSFWorkbook(fileStream);  //xls数据读入workbook  
            }
            //word文档
            XWPFDocument doc = new XWPFDocument(); //文档

            ISheet sheet = workbook.GetSheetAt(0);  //获取第一个工作表  
            IRow row;// = sheet.GetRow(0);            //新建当前工作表行数据  
            for (int i = 1; i <= sheet.LastRowNum; i++)  //对工作表每一行  
            {
                row = sheet.GetRow(i);   //row读入第i行数据  
                if (row != null)
                {
                    //标题
                    string title = row.GetCell(0).ToString();
                    Console.WriteLine(title);

                    XWPFParagraph p1 = doc.CreateParagraph();
                    p1.Alignment = ParagraphAlignment.LEFT;

                    XWPFRun r1 = p1.CreateRun();
                    r1.SetText(title);

                    //图片
                    string pictureaName = row.GetCell(1).ToString();
                    Console.WriteLine(pictureaName);

                    XWPFParagraph p2 = doc.CreateParagraph();
                    p2.Alignment = ParagraphAlignment.LEFT;

                    FileStream image = new FileStream(imagePath + "/" + pictureaName, FileMode.Open, FileAccess.Read);

                    XWPFRun r2 = p2.CreateRun();
                    r2.AddPicture(image, 5, imagePath + "/" + pictureaName, Units.ToEMU(500), Units.ToEMU(500));
                    r2.AddBreak();
                    if (i == sheet.LastRowNum)
                    {
                        last = true;
                    }
                }
                double d = (double) ((double)i / (double)sheet.LastRowNum) * 100;
                process = (int)Math.Floor(Convert.ToDouble(d));
  
                Console.WriteLine(process);

                this.Invoke(this.myIncrease, new object[] { process, last });


                //this.progressBar1.BeginInvoke(new EventHandler((sender, e) =>
                //{
                //    this.progressBar1.Value = process;
                //}), null);
            }
            Console.ReadLine();
            fileStream.Close();
            workbook.Close();

            string tradeTime = DateTime.Now.ToString("yyyyMMddHHmmss", DateTimeFormatInfo.InvariantInfo);
            FileStream out1 = null;
            try
            {
                out1 = new FileStream(@outPath +"/" + tradeTime + ".docx", FileMode.Create);
            }
            catch (Exception exception)
            {
                MessageBox.Show("出现错误：" + exception.Message);
                return;
            }

            doc.Write(out1);
            out1.Close();
        }

        private void ShowProcessBar()
        {
            myProcessBar = new Form2();
            myProcessBar.StartPosition = FormStartPosition.CenterScreen;

            myIncrease = new IncreaseHandle(myProcessBar.setValue);
            myProcessBar.ShowDialog();
            myProcessBar = null;
        }

        public static void stopThread()
        {
            thread.Abort();
        }

 
    }
}
