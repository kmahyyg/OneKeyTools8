namespace OneKeyTools
{
    partial class OK_Command
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
            this.关闭ToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.分隔符ToolStripMenuItem = new System.Windows.Forms.ToolStripSeparator();
            this.静音ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.说明ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.OKCommand = new System.Windows.Forms.Button();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.关闭ToolStripMenuItem1,
            this.分隔符ToolStripMenuItem,
            this.静音ToolStripMenuItem,
            this.说明ToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(153, 98);
            // 
            // 关闭ToolStripMenuItem1
            // 
            this.关闭ToolStripMenuItem1.Name = "关闭ToolStripMenuItem1";
            this.关闭ToolStripMenuItem1.Size = new System.Drawing.Size(152, 22);
            this.关闭ToolStripMenuItem1.Text = "关闭";
            this.关闭ToolStripMenuItem1.Click += new System.EventHandler(this.关闭ToolStripMenuItem1_Click);
            // 
            // 分隔符ToolStripMenuItem
            // 
            this.分隔符ToolStripMenuItem.Name = "分隔符ToolStripMenuItem";
            this.分隔符ToolStripMenuItem.Size = new System.Drawing.Size(149, 6);
            // 
            // 静音ToolStripMenuItem
            // 
            this.静音ToolStripMenuItem.Name = "静音ToolStripMenuItem";
            this.静音ToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.静音ToolStripMenuItem.Text = "静音";
            this.静音ToolStripMenuItem.Click += new System.EventHandler(this.静音ToolStripMenuItem_Click);
            // 
            // 说明ToolStripMenuItem
            // 
            this.说明ToolStripMenuItem.Name = "说明ToolStripMenuItem";
            this.说明ToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.说明ToolStripMenuItem.Text = "说明";
            this.说明ToolStripMenuItem.Click += new System.EventHandler(this.说明ToolStripMenuItem_Click);
            // 
            // OKCommand
            // 
            this.OKCommand.ContextMenuStrip = this.contextMenuStrip1;
            this.OKCommand.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.OKCommand.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.OKCommand.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.OKCommand.ForeColor = System.Drawing.Color.White;
            this.OKCommand.Location = new System.Drawing.Point(26, 12);
            this.OKCommand.Name = "OKCommand";
            this.OKCommand.Size = new System.Drawing.Size(122, 54);
            this.OKCommand.TabIndex = 2;
            this.OKCommand.Text = "单击此处";
            this.OKCommand.UseVisualStyleBackColor = true;
            this.OKCommand.Click += new System.EventHandler(this.OKCommand_Click);
            // 
            // OK_Command
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(67)))), ((int)(((byte)(67)))));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ClientSize = new System.Drawing.Size(172, 80);
            this.ContextMenuStrip = this.contextMenuStrip1;
            this.Controls.Add(this.OKCommand);
            this.DoubleBuffered = true;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "OK_Command";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "OK命令";
            this.TopMost = true;
            this.TransparencyKey = System.Drawing.SystemColors.Control;
            this.Load += new System.EventHandler(this.OKRecognition2_Load);
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 关闭ToolStripMenuItem1;
        private System.Windows.Forms.Button OKCommand;
        private System.Windows.Forms.ToolStripMenuItem 说明ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 静音ToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator 分隔符ToolStripMenuItem;

    }
}