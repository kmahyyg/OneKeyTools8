namespace OneKeyTools
{
    partial class Time_Count
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.label1 = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.fontDialog1 = new System.Windows.Forms.FontDialog();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.重新开始ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.切换显示格式ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.分隔符ToolStripMenuItem = new System.Windows.Forms.ToolStripSeparator();
            this.设置时间ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.下方框中输入秒数ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.时间ToolStripMenuItem = new System.Windows.Forms.ToolStripTextBox();
            this.修改提示ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.修改文字后回车ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.提示ToolStripMenuItem = new System.Windows.Forms.ToolStripTextBox();
            this.分隔符ToolStripMenuItem3 = new System.Windows.Forms.ToolStripSeparator();
            this.声音设置ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.修改文字ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.文字颜色ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.文字格式ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.修改背景ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.背景色ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.背景图片ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.分隔符ToolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.保存设置ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.恢复默认设置ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.分隔符ToolStripMenuItem2 = new System.Windows.Forms.ToolStripSeparator();
            this.使用说明ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.关闭ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(10, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(24, 26);
            this.label1.TabIndex = 0;
            this.label1.Text = "0";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label1.MouseClick += new System.Windows.Forms.MouseEventHandler(this.label1_MouseClick);
            // 
            // timer1
            // 
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.重新开始ToolStripMenuItem,
            this.切换显示格式ToolStripMenuItem,
            this.分隔符ToolStripMenuItem,
            this.设置时间ToolStripMenuItem,
            this.修改提示ToolStripMenuItem,
            this.修改文字ToolStripMenuItem,
            this.修改背景ToolStripMenuItem,
            this.分隔符ToolStripMenuItem1,
            this.保存设置ToolStripMenuItem,
            this.恢复默认设置ToolStripMenuItem,
            this.分隔符ToolStripMenuItem2,
            this.使用说明ToolStripMenuItem,
            this.关闭ToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(149, 242);
            // 
            // 重新开始ToolStripMenuItem
            // 
            this.重新开始ToolStripMenuItem.Name = "重新开始ToolStripMenuItem";
            this.重新开始ToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.重新开始ToolStripMenuItem.Text = "重新开始";
            this.重新开始ToolStripMenuItem.Click += new System.EventHandler(this.重新开始ToolStripMenuItem_Click);
            // 
            // 切换显示格式ToolStripMenuItem
            // 
            this.切换显示格式ToolStripMenuItem.Name = "切换显示格式ToolStripMenuItem";
            this.切换显示格式ToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.切换显示格式ToolStripMenuItem.Text = "切换显示格式";
            this.切换显示格式ToolStripMenuItem.Click += new System.EventHandler(this.切换显示格式ToolStripMenuItem_Click);
            // 
            // 分隔符ToolStripMenuItem
            // 
            this.分隔符ToolStripMenuItem.Name = "分隔符ToolStripMenuItem";
            this.分隔符ToolStripMenuItem.Size = new System.Drawing.Size(145, 6);
            // 
            // 设置时间ToolStripMenuItem
            // 
            this.设置时间ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.下方框中输入秒数ToolStripMenuItem,
            this.时间ToolStripMenuItem});
            this.设置时间ToolStripMenuItem.Name = "设置时间ToolStripMenuItem";
            this.设置时间ToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.设置时间ToolStripMenuItem.Text = "设置时间";
            // 
            // 下方框中输入秒数ToolStripMenuItem
            // 
            this.下方框中输入秒数ToolStripMenuItem.Name = "下方框中输入秒数ToolStripMenuItem";
            this.下方框中输入秒数ToolStripMenuItem.Size = new System.Drawing.Size(232, 22);
            this.下方框中输入秒数ToolStripMenuItem.Text = "输入时间(秒)后回车";
            // 
            // 时间ToolStripMenuItem
            // 
            this.时间ToolStripMenuItem.Name = "时间ToolStripMenuItem";
            this.时间ToolStripMenuItem.Size = new System.Drawing.Size(172, 23);
            this.时间ToolStripMenuItem.Text = "60";
            this.时间ToolStripMenuItem.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.时间ToolStripMenuItem_KeyPress);
            // 
            // 修改提示ToolStripMenuItem
            // 
            this.修改提示ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.修改文字后回车ToolStripMenuItem,
            this.提示ToolStripMenuItem,
            this.分隔符ToolStripMenuItem3,
            this.声音设置ToolStripMenuItem});
            this.修改提示ToolStripMenuItem.Name = "修改提示ToolStripMenuItem";
            this.修改提示ToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.修改提示ToolStripMenuItem.Text = "设置提示";
            // 
            // 修改文字后回车ToolStripMenuItem
            // 
            this.修改文字后回车ToolStripMenuItem.Name = "修改文字后回车ToolStripMenuItem";
            this.修改文字后回车ToolStripMenuItem.Size = new System.Drawing.Size(220, 22);
            this.修改文字后回车ToolStripMenuItem.Text = "输入文字后回车";
            // 
            // 提示ToolStripMenuItem
            // 
            this.提示ToolStripMenuItem.Name = "提示ToolStripMenuItem";
            this.提示ToolStripMenuItem.Size = new System.Drawing.Size(160, 23);
            this.提示ToolStripMenuItem.Text = "时间到";
            this.提示ToolStripMenuItem.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.提示ToolStripMenuItem_KeyPress);
            // 
            // 分隔符ToolStripMenuItem3
            // 
            this.分隔符ToolStripMenuItem3.Name = "分隔符ToolStripMenuItem3";
            this.分隔符ToolStripMenuItem3.Size = new System.Drawing.Size(217, 6);
            // 
            // 声音设置ToolStripMenuItem
            // 
            this.声音设置ToolStripMenuItem.Name = "声音设置ToolStripMenuItem";
            this.声音设置ToolStripMenuItem.Size = new System.Drawing.Size(220, 22);
            this.声音设置ToolStripMenuItem.Text = "设置声音";
            this.声音设置ToolStripMenuItem.Click += new System.EventHandler(this.声音设置ToolStripMenuItem_Click);
            // 
            // 修改文字ToolStripMenuItem
            // 
            this.修改文字ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.文字颜色ToolStripMenuItem,
            this.文字格式ToolStripMenuItem});
            this.修改文字ToolStripMenuItem.Name = "修改文字ToolStripMenuItem";
            this.修改文字ToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.修改文字ToolStripMenuItem.Text = "修改文字";
            // 
            // 文字颜色ToolStripMenuItem
            // 
            this.文字颜色ToolStripMenuItem.Name = "文字颜色ToolStripMenuItem";
            this.文字颜色ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.文字颜色ToolStripMenuItem.Text = "文字颜色";
            this.文字颜色ToolStripMenuItem.Click += new System.EventHandler(this.文字颜色ToolStripMenuItem_Click);
            // 
            // 文字格式ToolStripMenuItem
            // 
            this.文字格式ToolStripMenuItem.Name = "文字格式ToolStripMenuItem";
            this.文字格式ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.文字格式ToolStripMenuItem.Text = "文字格式";
            this.文字格式ToolStripMenuItem.Click += new System.EventHandler(this.文字格式ToolStripMenuItem_Click);
            // 
            // 修改背景ToolStripMenuItem
            // 
            this.修改背景ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.背景色ToolStripMenuItem,
            this.背景图片ToolStripMenuItem});
            this.修改背景ToolStripMenuItem.Name = "修改背景ToolStripMenuItem";
            this.修改背景ToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.修改背景ToolStripMenuItem.Text = "修改背景";
            // 
            // 背景色ToolStripMenuItem
            // 
            this.背景色ToolStripMenuItem.Name = "背景色ToolStripMenuItem";
            this.背景色ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.背景色ToolStripMenuItem.Text = "背景色";
            this.背景色ToolStripMenuItem.Click += new System.EventHandler(this.背景色ToolStripMenuItem_Click);
            // 
            // 背景图片ToolStripMenuItem
            // 
            this.背景图片ToolStripMenuItem.Name = "背景图片ToolStripMenuItem";
            this.背景图片ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.背景图片ToolStripMenuItem.Text = "背景图片";
            this.背景图片ToolStripMenuItem.Click += new System.EventHandler(this.背景图片ToolStripMenuItem_Click);
            // 
            // 分隔符ToolStripMenuItem1
            // 
            this.分隔符ToolStripMenuItem1.Name = "分隔符ToolStripMenuItem1";
            this.分隔符ToolStripMenuItem1.Size = new System.Drawing.Size(145, 6);
            // 
            // 保存设置ToolStripMenuItem
            // 
            this.保存设置ToolStripMenuItem.Name = "保存设置ToolStripMenuItem";
            this.保存设置ToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.保存设置ToolStripMenuItem.Text = "保存设置";
            this.保存设置ToolStripMenuItem.Click += new System.EventHandler(this.保存设置ToolStripMenuItem_Click);
            // 
            // 恢复默认设置ToolStripMenuItem
            // 
            this.恢复默认设置ToolStripMenuItem.Name = "恢复默认设置ToolStripMenuItem";
            this.恢复默认设置ToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.恢复默认设置ToolStripMenuItem.Text = "恢复默认";
            this.恢复默认设置ToolStripMenuItem.Click += new System.EventHandler(this.恢复默认设置ToolStripMenuItem_Click);
            // 
            // 分隔符ToolStripMenuItem2
            // 
            this.分隔符ToolStripMenuItem2.Name = "分隔符ToolStripMenuItem2";
            this.分隔符ToolStripMenuItem2.Size = new System.Drawing.Size(145, 6);
            // 
            // 使用说明ToolStripMenuItem
            // 
            this.使用说明ToolStripMenuItem.Name = "使用说明ToolStripMenuItem";
            this.使用说明ToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.使用说明ToolStripMenuItem.Text = "使用说明";
            this.使用说明ToolStripMenuItem.Click += new System.EventHandler(this.使用说明ToolStripMenuItem_Click);
            // 
            // 关闭ToolStripMenuItem
            // 
            this.关闭ToolStripMenuItem.Name = "关闭ToolStripMenuItem";
            this.关闭ToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.关闭ToolStripMenuItem.Text = "关闭";
            this.关闭ToolStripMenuItem.Click += new System.EventHandler(this.关闭ToolStripMenuItem_Click);
            // 
            // Time_Count
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(67)))), ((int)(((byte)(67)))));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(300, 52);
            this.ContextMenuStrip = this.contextMenuStrip1;
            this.ControlBox = false;
            this.Controls.Add(this.label1);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Time_Count";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "定时器";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.Time_Count_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Time_Count_KeyDown);
            this.Resize += new System.EventHandler(this.Time_Count_Resize);
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.FontDialog fontDialog1;
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripSeparator 分隔符ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 修改文字ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 文字颜色ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 文字格式ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 修改背景ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 背景色ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 背景图片ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 设置时间ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 下方框中输入秒数ToolStripMenuItem;
        private System.Windows.Forms.ToolStripTextBox 时间ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 使用说明ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 关闭ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 重新开始ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 切换显示格式ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 保存设置ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 修改提示ToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator 分隔符ToolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem 修改文字后回车ToolStripMenuItem;
        private System.Windows.Forms.ToolStripTextBox 提示ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 恢复默认设置ToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator 分隔符ToolStripMenuItem1;
        private System.Windows.Forms.ToolStripSeparator 分隔符ToolStripMenuItem3;
        private System.Windows.Forms.ToolStripMenuItem 声音设置ToolStripMenuItem;
    }
}