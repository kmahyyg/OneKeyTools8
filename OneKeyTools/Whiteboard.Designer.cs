namespace OneKeyTools
{
    partial class Whiteboard
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
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.新建ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.复制ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.设置宽高ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripTextBox();
            this.载入ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.设置颜色ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.取色器ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.颜色框ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.透明度ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.保存设置ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.恢复默认ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.关闭ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.DropShadowEnabled = false;
            this.contextMenuStrip1.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.新建ToolStripMenuItem,
            this.复制ToolStripMenuItem,
            this.toolStripSeparator2,
            this.设置宽高ToolStripMenuItem,
            this.载入ToolStripMenuItem,
            this.设置颜色ToolStripMenuItem,
            this.透明度ToolStripMenuItem,
            this.toolStripSeparator1,
            this.保存设置ToolStripMenuItem,
            this.恢复默认ToolStripMenuItem,
            this.toolStripSeparator3,
            this.关闭ToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(137, 220);
            // 
            // 新建ToolStripMenuItem
            // 
            this.新建ToolStripMenuItem.Name = "新建ToolStripMenuItem";
            this.新建ToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.新建ToolStripMenuItem.Text = "新建";
            this.新建ToolStripMenuItem.TextDirection = System.Windows.Forms.ToolStripTextDirection.Horizontal;
            this.新建ToolStripMenuItem.Visible = false;
            this.新建ToolStripMenuItem.Click += new System.EventHandler(this.新建ToolStripMenuItem_Click);
            // 
            // 复制ToolStripMenuItem
            // 
            this.复制ToolStripMenuItem.Name = "复制ToolStripMenuItem";
            this.复制ToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.复制ToolStripMenuItem.Text = "复制";
            this.复制ToolStripMenuItem.Click += new System.EventHandler(this.复制ToolStripMenuItem_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(133, 6);
            // 
            // 设置宽高ToolStripMenuItem
            // 
            this.设置宽高ToolStripMenuItem.Checked = true;
            this.设置宽高ToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked;
            this.设置宽高ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem2});
            this.设置宽高ToolStripMenuItem.Name = "设置宽高ToolStripMenuItem";
            this.设置宽高ToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.设置宽高ToolStripMenuItem.Text = "设置宽高";
            this.设置宽高ToolStripMenuItem.Click += new System.EventHandler(this.设置宽高ToolStripMenuItem_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(152, 23);
            this.toolStripMenuItem2.Text = "200,200";
            this.toolStripMenuItem2.TextChanged += new System.EventHandler(this.toolStripMenuItem2_TextChanged);
            // 
            // 载入ToolStripMenuItem
            // 
            this.载入ToolStripMenuItem.Name = "载入ToolStripMenuItem";
            this.载入ToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.载入ToolStripMenuItem.Text = "载入图像";
            this.载入ToolStripMenuItem.Click += new System.EventHandler(this.载入ToolStripMenuItem_Click);
            // 
            // 设置颜色ToolStripMenuItem
            // 
            this.设置颜色ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.取色器ToolStripMenuItem,
            this.颜色框ToolStripMenuItem});
            this.设置颜色ToolStripMenuItem.Name = "设置颜色ToolStripMenuItem";
            this.设置颜色ToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.设置颜色ToolStripMenuItem.Text = "设置颜色";
            // 
            // 取色器ToolStripMenuItem
            // 
            this.取色器ToolStripMenuItem.Name = "取色器ToolStripMenuItem";
            this.取色器ToolStripMenuItem.Size = new System.Drawing.Size(112, 22);
            this.取色器ToolStripMenuItem.Text = "取色器";
            this.取色器ToolStripMenuItem.Click += new System.EventHandler(this.取色器ToolStripMenuItem_Click);
            // 
            // 颜色框ToolStripMenuItem
            // 
            this.颜色框ToolStripMenuItem.Name = "颜色框ToolStripMenuItem";
            this.颜色框ToolStripMenuItem.Size = new System.Drawing.Size(112, 22);
            this.颜色框ToolStripMenuItem.Text = "颜色框";
            this.颜色框ToolStripMenuItem.Click += new System.EventHandler(this.颜色框ToolStripMenuItem_Click);
            // 
            // 透明度ToolStripMenuItem
            // 
            this.透明度ToolStripMenuItem.Name = "透明度ToolStripMenuItem";
            this.透明度ToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.透明度ToolStripMenuItem.Text = "设置透明度";
            this.透明度ToolStripMenuItem.Click += new System.EventHandler(this.透明度ToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(133, 6);
            // 
            // 保存设置ToolStripMenuItem
            // 
            this.保存设置ToolStripMenuItem.Name = "保存设置ToolStripMenuItem";
            this.保存设置ToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.保存设置ToolStripMenuItem.Text = "保存设置";
            this.保存设置ToolStripMenuItem.Click += new System.EventHandler(this.保存设置ToolStripMenuItem_Click);
            // 
            // 恢复默认ToolStripMenuItem
            // 
            this.恢复默认ToolStripMenuItem.Name = "恢复默认ToolStripMenuItem";
            this.恢复默认ToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.恢复默认ToolStripMenuItem.Text = "恢复默认";
            this.恢复默认ToolStripMenuItem.Click += new System.EventHandler(this.恢复默认ToolStripMenuItem_Click);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(133, 6);
            // 
            // 关闭ToolStripMenuItem
            // 
            this.关闭ToolStripMenuItem.Name = "关闭ToolStripMenuItem";
            this.关闭ToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.关闭ToolStripMenuItem.Text = "关闭界面";
            this.关闭ToolStripMenuItem.Click += new System.EventHandler(this.关闭ToolStripMenuItem_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // Whiteboard
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(248)))), ((int)(((byte)(248)))));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(200, 200);
            this.ContextMenuStrip = this.contextMenuStrip1;
            this.Cursor = System.Windows.Forms.Cursors.SizeAll;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Whiteboard";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "演示白板";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.gif3_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.gif3_KeyDown);
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 载入ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 关闭ToolStripMenuItem;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.ToolStripMenuItem 设置宽高ToolStripMenuItem;
        private System.Windows.Forms.ToolStripTextBox toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem 设置颜色ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 取色器ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 颜色框ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 透明度ToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem 保存设置ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 恢复默认ToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripMenuItem 复制ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 新建ToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
    }
}