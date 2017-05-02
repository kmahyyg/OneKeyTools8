namespace OneKeyTools
{
    partial class Time_Clock
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
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.fontDialog1 = new System.Windows.Forms.FontDialog();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.切换时钟格式ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.分隔符ToolStripMenuItem = new System.Windows.Forms.ToolStripSeparator();
            this.修改文字格式ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.文字颜色ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.文字格式ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.修改背景ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.背景色ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.背景图片ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.关闭ToolStripMenuItem = new System.Windows.Forms.ToolStripSeparator();
            this.保存设置ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.恢复默认ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.分隔符ToolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.使用说明ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.关闭ToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Cursor = System.Windows.Forms.Cursors.Default;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(8, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(117, 26);
            this.label1.TabIndex = 0;
            this.label1.Text = "OK数字时钟";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // colorDialog1
            // 
            this.colorDialog1.AnyColor = true;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.切换时钟格式ToolStripMenuItem,
            this.分隔符ToolStripMenuItem,
            this.修改文字格式ToolStripMenuItem,
            this.修改背景ToolStripMenuItem,
            this.关闭ToolStripMenuItem,
            this.保存设置ToolStripMenuItem,
            this.恢复默认ToolStripMenuItem,
            this.分隔符ToolStripMenuItem1,
            this.使用说明ToolStripMenuItem,
            this.关闭ToolStripMenuItem1});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(149, 176);
            // 
            // 切换时钟格式ToolStripMenuItem
            // 
            this.切换时钟格式ToolStripMenuItem.Name = "切换时钟格式ToolStripMenuItem";
            this.切换时钟格式ToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.切换时钟格式ToolStripMenuItem.Text = "切换显示格式";
            this.切换时钟格式ToolStripMenuItem.Click += new System.EventHandler(this.切换时钟格式ToolStripMenuItem_Click);
            // 
            // 分隔符ToolStripMenuItem
            // 
            this.分隔符ToolStripMenuItem.Name = "分隔符ToolStripMenuItem";
            this.分隔符ToolStripMenuItem.Size = new System.Drawing.Size(145, 6);
            // 
            // 修改文字格式ToolStripMenuItem
            // 
            this.修改文字格式ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.文字颜色ToolStripMenuItem,
            this.文字格式ToolStripMenuItem});
            this.修改文字格式ToolStripMenuItem.Name = "修改文字格式ToolStripMenuItem";
            this.修改文字格式ToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.修改文字格式ToolStripMenuItem.Text = "修改文字";
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
            // 关闭ToolStripMenuItem
            // 
            this.关闭ToolStripMenuItem.Name = "关闭ToolStripMenuItem";
            this.关闭ToolStripMenuItem.Size = new System.Drawing.Size(145, 6);
            // 
            // 保存设置ToolStripMenuItem
            // 
            this.保存设置ToolStripMenuItem.Name = "保存设置ToolStripMenuItem";
            this.保存设置ToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.保存设置ToolStripMenuItem.Text = "保存设置";
            this.保存设置ToolStripMenuItem.Click += new System.EventHandler(this.保存设置ToolStripMenuItem_Click);
            // 
            // 恢复默认ToolStripMenuItem
            // 
            this.恢复默认ToolStripMenuItem.Name = "恢复默认ToolStripMenuItem";
            this.恢复默认ToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.恢复默认ToolStripMenuItem.Text = "恢复默认";
            this.恢复默认ToolStripMenuItem.Click += new System.EventHandler(this.恢复默认ToolStripMenuItem_Click);
            // 
            // 分隔符ToolStripMenuItem1
            // 
            this.分隔符ToolStripMenuItem1.Name = "分隔符ToolStripMenuItem1";
            this.分隔符ToolStripMenuItem1.Size = new System.Drawing.Size(145, 6);
            // 
            // 使用说明ToolStripMenuItem
            // 
            this.使用说明ToolStripMenuItem.Name = "使用说明ToolStripMenuItem";
            this.使用说明ToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.使用说明ToolStripMenuItem.Text = "使用说明";
            this.使用说明ToolStripMenuItem.Click += new System.EventHandler(this.使用说明ToolStripMenuItem_Click);
            // 
            // 关闭ToolStripMenuItem1
            // 
            this.关闭ToolStripMenuItem1.Name = "关闭ToolStripMenuItem1";
            this.关闭ToolStripMenuItem1.Size = new System.Drawing.Size(148, 22);
            this.关闭ToolStripMenuItem1.Text = "关闭";
            this.关闭ToolStripMenuItem1.Click += new System.EventHandler(this.关闭ToolStripMenuItem1_Click);
            // 
            // Time_Clock
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(67)))), ((int)(((byte)(67)))));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(300, 52);
            this.ContextMenuStrip = this.contextMenuStrip1;
            this.ControlBox = false;
            this.Controls.Add(this.label1);
            this.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.ForeColor = System.Drawing.Color.Black;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Time_Clock";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "数字时钟";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.time1_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.time1_KeyDown);
            this.Resize += new System.EventHandler(this.time1_Resize);
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.FontDialog fontDialog1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 修改文字格式ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 文字颜色ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 文字格式ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 修改背景ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 背景色ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 背景图片ToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator 关闭ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 关闭ToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem 切换时钟格式ToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator 分隔符ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 使用说明ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 保存设置ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 恢复默认ToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator 分隔符ToolStripMenuItem1;
    }
}