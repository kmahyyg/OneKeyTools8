namespace OneKeyTools
{
    partial class Audio_Record
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
            this.TitleLabel = new System.Windows.Forms.Label();
            this.CloseButton = new System.Windows.Forms.Button();
            this.BeginRecognition = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.RecordAudiosBox = new System.Windows.Forms.ListBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.删除所选ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.删除所有ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // TitleLabel
            // 
            this.TitleLabel.AutoSize = true;
            this.TitleLabel.BackColor = System.Drawing.Color.Transparent;
            this.TitleLabel.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.TitleLabel.Location = new System.Drawing.Point(54, 5);
            this.TitleLabel.Name = "TitleLabel";
            this.TitleLabel.Size = new System.Drawing.Size(65, 20);
            this.TitleLabel.TabIndex = 0;
            this.TitleLabel.Text = "录音工具";
            this.TitleLabel.Click += new System.EventHandler(this.TitleLabel_Click);
            // 
            // CloseButton
            // 
            this.CloseButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(67)))), ((int)(((byte)(67)))));
            this.CloseButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.CloseButton.FlatAppearance.BorderSize = 0;
            this.CloseButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.CloseButton.Font = new System.Drawing.Font("微软雅黑", 8F);
            this.CloseButton.ForeColor = System.Drawing.Color.White;
            this.CloseButton.Location = new System.Drawing.Point(10, 0);
            this.CloseButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.CloseButton.Name = "CloseButton";
            this.CloseButton.Size = new System.Drawing.Size(40, 27);
            this.CloseButton.TabIndex = 1;
            this.CloseButton.Text = "关闭";
            this.CloseButton.UseVisualStyleBackColor = false;
            this.CloseButton.Click += new System.EventHandler(this.CloseButton_Click);
            // 
            // BeginRecognition
            // 
            this.BeginRecognition.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(67)))), ((int)(((byte)(67)))));
            this.BeginRecognition.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.BeginRecognition.Cursor = System.Windows.Forms.Cursors.Hand;
            this.BeginRecognition.FlatAppearance.BorderSize = 0;
            this.BeginRecognition.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BeginRecognition.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.BeginRecognition.ForeColor = System.Drawing.Color.White;
            this.BeginRecognition.Location = new System.Drawing.Point(97, 163);
            this.BeginRecognition.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.BeginRecognition.Name = "BeginRecognition";
            this.BeginRecognition.Size = new System.Drawing.Size(101, 34);
            this.BeginRecognition.TabIndex = 2;
            this.BeginRecognition.Tag = "";
            this.BeginRecognition.Text = "开始";
            this.toolTip1.SetToolTip(this.BeginRecognition, "开始/停止录音");
            this.BeginRecognition.UseVisualStyleBackColor = false;
            this.BeginRecognition.Click += new System.EventHandler(this.BeginRecognition_Click);
            // 
            // toolTip1
            // 
            this.toolTip1.AutomaticDelay = 100;
            this.toolTip1.AutoPopDelay = 3000;
            this.toolTip1.InitialDelay = 100;
            this.toolTip1.ReshowDelay = 20;
            // 
            // RecordAudiosBox
            // 
            this.RecordAudiosBox.ContextMenuStrip = this.contextMenuStrip1;
            this.RecordAudiosBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.RecordAudiosBox.FormattingEnabled = true;
            this.RecordAudiosBox.ItemHeight = 17;
            this.RecordAudiosBox.Location = new System.Drawing.Point(10, 33);
            this.RecordAudiosBox.Name = "RecordAudiosBox";
            this.RecordAudiosBox.Size = new System.Drawing.Size(275, 123);
            this.RecordAudiosBox.TabIndex = 3;
            this.RecordAudiosBox.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.RecordAudiosBox_MouseDoubleClick);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.删除所选ToolStripMenuItem,
            this.删除所有ToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(125, 48);
            // 
            // 删除所选ToolStripMenuItem
            // 
            this.删除所选ToolStripMenuItem.Name = "删除所选ToolStripMenuItem";
            this.删除所选ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.删除所选ToolStripMenuItem.Text = "删除所选";
            this.删除所选ToolStripMenuItem.Click += new System.EventHandler(this.删除所选ToolStripMenuItem_Click);
            // 
            // 删除所有ToolStripMenuItem
            // 
            this.删除所有ToolStripMenuItem.Name = "删除所有ToolStripMenuItem";
            this.删除所有ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.删除所有ToolStripMenuItem.Text = "删除所有";
            this.删除所有ToolStripMenuItem.Click += new System.EventHandler(this.删除所有ToolStripMenuItem_Click);
            // 
            // Audio_Record
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(248)))), ((int)(((byte)(248)))));
            this.ClientSize = new System.Drawing.Size(295, 218);
            this.Controls.Add(this.RecordAudiosBox);
            this.Controls.Add(this.BeginRecognition);
            this.Controls.Add(this.TitleLabel);
            this.Controls.Add(this.CloseButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Audio_Record";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "录音工具";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.Audio_Record_Load);
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label TitleLabel;
        private System.Windows.Forms.Button CloseButton;
        private System.Windows.Forms.Button BeginRecognition;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.ListBox RecordAudiosBox;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 删除所选ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 删除所有ToolStripMenuItem;
    }
}