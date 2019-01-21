using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace excelTool
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        public delegate void MyDelegate();

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        public void setValue(int value, bool last)
        {
            this.progressBar1.Value = value;
            this.label1.Text = value.ToString() + "%";
            if (last == true) {
                this.Close();
                MessageBox.Show("任务已完成！");
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBoxButtons messButton = MessageBoxButtons.OKCancel;
            DialogResult dr = MessageBox.Show("确定要取消吗?", "提示", messButton);
            if (dr == DialogResult.OK)//如果点击“确定”按钮
            {
                MyDelegate myDelegate = new MyDelegate(Form1.stopThread);
                myDelegate();
                this.Close();
            }
        }
    }
}
