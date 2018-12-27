namespace JingWuTong

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
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.gpsTimer = new System.Windows.Forms.Timer(this.components);
            this.showTimer = new System.Windows.Forms.Timer(this.components);
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.系统设置ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ycsz_ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.bdsz_ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnRecive = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // gpsTimer
            // 
            this.gpsTimer.Interval = 5000;
            this.gpsTimer.Tick += new System.EventHandler(this.gpsTimer_Tick);
            // 
            // showTimer
            // 
            this.showTimer.Interval = 5000;
            this.showTimer.Tick += new System.EventHandler(this.showTimer_Tick);
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("微软雅黑", 11F);
            this.textBox1.Location = new System.Drawing.Point(105, 62);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(430, 166);
            this.textBox1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(12, 62);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 19);
            this.label1.TabIndex = 1;
            this.label1.Text = "接收状况：";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.系统设置ToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(568, 29);
            this.menuStrip1.TabIndex = 2;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // 系统设置ToolStripMenuItem
            // 
            this.系统设置ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ycsz_ToolStripMenuItem,
            this.bdsz_ToolStripMenuItem});
            this.系统设置ToolStripMenuItem.Font = new System.Drawing.Font("微软雅黑", 12F);
            this.系统设置ToolStripMenuItem.Name = "系统设置ToolStripMenuItem";
            this.系统设置ToolStripMenuItem.Size = new System.Drawing.Size(86, 25);
            this.系统设置ToolStripMenuItem.Text = "系统设置";
            // 
            // ycsz_ToolStripMenuItem
            // 
            this.ycsz_ToolStripMenuItem.Name = "ycsz_ToolStripMenuItem";
            this.ycsz_ToolStripMenuItem.Size = new System.Drawing.Size(144, 26);
            this.ycsz_ToolStripMenuItem.Text = "远程设置";
            this.ycsz_ToolStripMenuItem.Click += new System.EventHandler(this.ycsz_ToolStripMenuItem_Click);
            // 
            // bdsz_ToolStripMenuItem
            // 
            this.bdsz_ToolStripMenuItem.Name = "bdsz_ToolStripMenuItem";
            this.bdsz_ToolStripMenuItem.Size = new System.Drawing.Size(144, 26);
            this.bdsz_ToolStripMenuItem.Text = "本地设置";
            // 
            // btnRecive
            // 
            this.btnRecive.Font = new System.Drawing.Font("微软雅黑", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnRecive.Location = new System.Drawing.Point(105, 235);
            this.btnRecive.Name = "btnRecive";
            this.btnRecive.Size = new System.Drawing.Size(125, 28);
            this.btnRecive.TabIndex = 3;
            this.btnRecive.Text = "开始接收";
            this.btnRecive.UseVisualStyleBackColor = true;
            this.btnRecive.Click += new System.EventHandler(this.btnRecive_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(568, 312);
            this.Controls.Add(this.btnRecive);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "数据接收";
            this.TransparencyKey = System.Drawing.Color.LightGray;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Timer gpsTimer;
        private System.Windows.Forms.Timer showTimer;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 系统设置ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ycsz_ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem bdsz_ToolStripMenuItem;
        private System.Windows.Forms.Button btnRecive;
    }
}

