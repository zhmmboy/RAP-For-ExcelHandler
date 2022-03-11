using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelHandler
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            //this.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            //this.panel2.Dock = DockStyle.Top;
            //this.dataGridView1.Dock = DockStyle.Fill;

        }

        private void btnLoadFile_Click(object sender, EventArgs e)
        {

            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            folderBrowser.SelectedPath = "";
            folderBrowser.Description = "请选择要处理Excel的文件目录";
            //folderBrowser.ShowNewFolderButton = true;
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                this.textBox1.Text = folderBrowser.SelectedPath;
            }
            else
            {
                return;
            }


            //openFileDialog1.Title = "请选择要处理的Excel文件夹。";
            //openFileDialog1.Filter = "所有文件(*.*)|*.*";
            //if (openFileDialog1.ShowDialog() == DialogResult.OK)
            //{
            //    textBox1.Text = openFileDialog1.FileName;
            //}

            var fileDic = textBox1.Text;
            List<dynamic> lstD = new List<dynamic>();


            List<string> lstRt = new List<string>();
            GetFiles(lstRt, fileDic);
            var _index = 0;

            DataTable dt = new DataTable();
            dt.Columns.Add("_Index");
            dt.Columns.Add("Path");
            dt.Columns.Add("Status");
            dt.Columns.Add("Msg");

            foreach (var path in lstRt)
            {
                _index++;
                DataRow dr = dt.NewRow();
                dr[0] = _index.ToString();
                dr[1] = path;
                dr[2] = "待处理";
                dt.Rows.Add(dr);
            }

            this.dataGridView1.DataSource = dt;

            lblMsg.Text = string.Format("共找到 {0} 个待处理的Excel文件。", dt.Rows.Count);
        }


        private void GetFiles(List<string> lstRt, string fileDic)
        {
            var arrPath = System.IO.Directory.GetFileSystemEntries(fileDic);
            foreach (var item in arrPath)
            {
                //文件夹
                if (System.IO.Directory.Exists(item))
                {
                    GetFiles(lstRt, item);
                }
                else
                {
                    var ext = System.IO.Path.GetExtension(item);
                    if (ext.ToLower() == ".xls" || ext.ToLower() == ".xlsx" || ext.ToLower() == ".xlsm")
                    {
                        //如果是文件
                        lstRt.Add(item);
                    }
                }
            }
        }

        private void btnHandler_Click(object sender, EventArgs e)
        {
            
            if (textBox1.Text.Trim()== "请选择文件所在目录")
            {
                MessageBox.Show("请选择文件所在目录");
                return;
            }
            if (string.IsNullOrWhiteSpace(textBox2.Text.Trim()) || textBox2.Text.Trim()== "请输入VBA函数名称，例：Sheet1.Test")
            {
                MessageBox.Show("请输入VBA函数名称，例：Sheet1.Test");
                return;
            }

            if (dataGridView1.DataSource != null)
            {
                var dt = dataGridView1.DataSource as DataTable;
                var total = 0;
                var totalError = 0;

                //循环处理文件
                foreach (DataRow item in dt.Rows)
                {
                    try
                    {
                        this.lblMsg.Text = string.Format("正在处理：{0}", item["Path"].ToString());

                        object obj = new object();
                        new ExcelHelper().RunExcelMacro(item["Path"].ToString(), textBox2.Text, null, out obj, true);

                        this.lblMsg.Text = string.Format("处理完成：{0}", item["Path"].ToString());

                        dataGridView1.Rows[total].Cells[2].Value = "已处理";
                        dataGridView1.Rows[total].Cells[2].Value = "处理完成";

                    }
                    catch (Exception ex)
                    {
                        totalError++;
                        dataGridView1.Rows[total].Cells[2].Value = "处理异常：" + ex.Message;
                    }

                    total++;
                }

                this.lblMsg.Text = string.Format("一共处理完成：{0} 个文件,正常：{1} 个，异常：{2} 个。", dt.Rows.Count.ToString(), (dt.Rows.Count - totalError).ToString(), totalError.ToString());
            }
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                MessageBox.Show("请选择文件夹"); return;
            }

            var fileDic = textBox1.Text;
            //判断文件目录是否存在
            if (!System.IO.Directory.Exists(fileDic))
            {
                MessageBox.Show("该目录不存在"); return;
            }


        }

        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (textBox2.Text.Trim() == "请输入VBA函数名称，例：Sheet1.Test")
            {
                textBox2.Text = string.Empty;
            }
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            if (textBox2.Text.Trim() == "请输入VBA函数名称，例：Sheet1.Test")
            {
                textBox2.Text = string.Empty;
            }
        }

        private void textBox2_MouseLeave(object sender, EventArgs e)
        {
            if (textBox2.Text.Trim() == "")
            {
                textBox2.Text = "请输入VBA函数名称，例：Sheet1.Test";
            }
        }
    }
}
