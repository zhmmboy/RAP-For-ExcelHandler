
namespace ExcelHandler
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.btnLoadFile = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.btnHandler = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.lblMsg = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.序号 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.文件目录 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.状态 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.说明 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnLoadFile
            // 
            this.btnLoadFile.Location = new System.Drawing.Point(232, 15);
            this.btnLoadFile.Name = "btnLoadFile";
            this.btnLoadFile.Size = new System.Drawing.Size(113, 39);
            this.btnLoadFile.TabIndex = 11;
            this.btnLoadFile.Text = "选择文件目录";
            this.btnLoadFile.UseVisualStyleBackColor = true;
            this.btnLoadFile.Click += new System.EventHandler(this.btnLoadFile_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(16, 16);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(210, 38);
            this.textBox1.TabIndex = 12;
            this.textBox1.Text = "请选择文件所在目录";
            // 
            // btnHandler
            // 
            this.btnHandler.Location = new System.Drawing.Point(609, 15);
            this.btnHandler.Name = "btnHandler";
            this.btnHandler.Size = new System.Drawing.Size(104, 39);
            this.btnHandler.TabIndex = 14;
            this.btnHandler.Text = "开始处理";
            this.btnHandler.UseVisualStyleBackColor = true;
            this.btnHandler.Click += new System.EventHandler(this.btnHandler_Click);
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.label1);
            this.panel2.Controls.Add(this.textBox2);
            this.panel2.Controls.Add(this.lblMsg);
            this.panel2.Controls.Add(this.btnHandler);
            this.panel2.Controls.Add(this.textBox1);
            this.panel2.Controls.Add(this.btnLoadFile);
            this.panel2.Location = new System.Drawing.Point(23, 14);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1189, 98);
            this.panel2.TabIndex = 16;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.OrangeRed;
            this.label1.Location = new System.Drawing.Point(12, 77);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(812, 15);
            this.label1.TabIndex = 17;
            this.label1.Text = "程序使用步骤：1.选择文件目录；2.输入VBA宏函数名称，例：Sheet1.Test；3.点击”开始处理“按钮，开始处理文件。";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(368, 16);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(235, 38);
            this.textBox2.TabIndex = 16;
            this.textBox2.Text = "请输入VBA函数名称，例：Sheet1.Test";
            this.textBox2.Click += new System.EventHandler(this.textBox2_Click);
            this.textBox2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox2_KeyDown);
            this.textBox2.MouseLeave += new System.EventHandler(this.textBox2_MouseLeave);
            // 
            // lblMsg
            // 
            this.lblMsg.AutoSize = true;
            this.lblMsg.Location = new System.Drawing.Point(719, 27);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(52, 15);
            this.lblMsg.TabIndex = 15;
            this.lblMsg.Text = "消息：";
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.序号,
            this.文件目录,
            this.状态,
            this.说明});
            this.dataGridView1.Location = new System.Drawing.Point(20, 118);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 27;
            this.dataGridView1.Size = new System.Drawing.Size(1200, 474);
            this.dataGridView1.TabIndex = 15;
            // 
            // 序号
            // 
            this.序号.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.序号.DataPropertyName = "_Index";
            this.序号.FillWeight = 6F;
            this.序号.HeaderText = "序号";
            this.序号.MinimumWidth = 6;
            this.序号.Name = "序号";
            this.序号.ReadOnly = true;
            // 
            // 文件目录
            // 
            this.文件目录.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.文件目录.DataPropertyName = "Path";
            this.文件目录.FillWeight = 28F;
            this.文件目录.HeaderText = "文件目录";
            this.文件目录.MinimumWidth = 6;
            this.文件目录.Name = "文件目录";
            this.文件目录.ReadOnly = true;
            // 
            // 状态
            // 
            this.状态.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.状态.DataPropertyName = "Status";
            this.状态.FillWeight = 6F;
            this.状态.HeaderText = "状态";
            this.状态.MinimumWidth = 6;
            this.状态.Name = "状态";
            this.状态.ReadOnly = true;
            // 
            // 说明
            // 
            this.说明.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.说明.DataPropertyName = "Msg";
            this.说明.FillWeight = 20F;
            this.说明.HeaderText = "说明";
            this.说明.MinimumWidth = 6;
            this.说明.Name = "说明";
            this.说明.ReadOnly = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(1233, 608);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.panel2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.Text = "欢迎使用 Excel 文件处理小助手！";
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button btnLoadFile;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button btnHandler;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label lblMsg;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.DataGridViewTextBoxColumn 序号;
        private System.Windows.Forms.DataGridViewTextBoxColumn 文件目录;
        private System.Windows.Forms.DataGridViewTextBoxColumn 状态;
        private System.Windows.Forms.DataGridViewTextBoxColumn 说明;
        private System.Windows.Forms.Label label1;
    }
}

