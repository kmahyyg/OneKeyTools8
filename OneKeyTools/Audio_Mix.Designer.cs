namespace OneKeyTools
{
    partial class Audio_Mix
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
            this.Total_Label = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.Current_Label = new System.Windows.Forms.Label();
            this.StopButton = new System.Windows.Forms.Button();
            this.PlayButton = new System.Windows.Forms.Button();
            this.trackBar1 = new System.Windows.Forms.TrackBar();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.AudioMix = new System.Windows.Forms.Button();
            this.AddAudioFile = new System.Windows.Forms.Button();
            this.AudioFilesBox = new System.Windows.Forms.ListBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.DeleteSelection = new System.Windows.Forms.ToolStripMenuItem();
            this.DeleteAll = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.trackBar1)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // TitleLabel
            // 
            this.TitleLabel.AutoSize = true;
            this.TitleLabel.BackColor = System.Drawing.Color.Transparent;
            this.TitleLabel.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.TitleLabel.Location = new System.Drawing.Point(54, 7);
            this.TitleLabel.Name = "TitleLabel";
            this.TitleLabel.Size = new System.Drawing.Size(56, 17);
            this.TitleLabel.TabIndex = 0;
            this.TitleLabel.Text = "音频混合";
            this.TitleLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
            this.toolTip1.SetToolTip(this.CloseButton, "关闭界面");
            this.CloseButton.UseVisualStyleBackColor = false;
            this.CloseButton.Click += new System.EventHandler(this.CloseButton_Click);
            // 
            // Total_Label
            // 
            this.Total_Label.AutoSize = true;
            this.Total_Label.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Total_Label.Location = new System.Drawing.Point(133, 213);
            this.Total_Label.Name = "Total_Label";
            this.Total_Label.Size = new System.Drawing.Size(56, 17);
            this.Total_Label.TabIndex = 0;
            this.Total_Label.Text = "00:00:00";
            this.toolTip1.SetToolTip(this.Total_Label, "总时间");
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(120, 213);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(13, 17);
            this.label2.TabIndex = 0;
            this.label2.Text = "/";
            // 
            // Current_Label
            // 
            this.Current_Label.AutoSize = true;
            this.Current_Label.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Current_Label.Location = new System.Drawing.Point(64, 213);
            this.Current_Label.Name = "Current_Label";
            this.Current_Label.Size = new System.Drawing.Size(56, 17);
            this.Current_Label.TabIndex = 0;
            this.Current_Label.Text = "00:00:00";
            this.toolTip1.SetToolTip(this.Current_Label, "当前时间");
            // 
            // StopButton
            // 
            this.StopButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.StopButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.StopButton.Enabled = false;
            this.StopButton.FlatAppearance.BorderSize = 0;
            this.StopButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.StopButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.StopButton.Location = new System.Drawing.Point(125, 240);
            this.StopButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.StopButton.Name = "StopButton";
            this.StopButton.Size = new System.Drawing.Size(64, 32);
            this.StopButton.TabIndex = 7;
            this.StopButton.Text = "停止";
            this.toolTip1.SetToolTip(this.StopButton, "停止播放音频");
            this.StopButton.UseVisualStyleBackColor = false;
            this.StopButton.Click += new System.EventHandler(this.StopButton_Click);
            // 
            // PlayButton
            // 
            this.PlayButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.PlayButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PlayButton.FlatAppearance.BorderSize = 0;
            this.PlayButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.PlayButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.PlayButton.Location = new System.Drawing.Point(54, 240);
            this.PlayButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.PlayButton.Name = "PlayButton";
            this.PlayButton.Size = new System.Drawing.Size(64, 32);
            this.PlayButton.TabIndex = 5;
            this.PlayButton.Text = "播放";
            this.toolTip1.SetToolTip(this.PlayButton, "播放音频");
            this.PlayButton.UseVisualStyleBackColor = false;
            this.PlayButton.Click += new System.EventHandler(this.PlayButton_Click);
            // 
            // trackBar1
            // 
            this.trackBar1.Cursor = System.Windows.Forms.Cursors.SizeWE;
            this.trackBar1.Enabled = false;
            this.trackBar1.Location = new System.Drawing.Point(20, 186);
            this.trackBar1.Maximum = 100;
            this.trackBar1.Name = "trackBar1";
            this.trackBar1.Size = new System.Drawing.Size(216, 45);
            this.trackBar1.TabIndex = 6;
            this.trackBar1.TickStyle = System.Windows.Forms.TickStyle.None;
            this.trackBar1.Scroll += new System.EventHandler(this.trackBar1_Scroll);
            this.trackBar1.ValueChanged += new System.EventHandler(this.trackBar1_ValueChanged);
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // toolTip1
            // 
            this.toolTip1.AutomaticDelay = 100;
            this.toolTip1.AutoPopDelay = 2000;
            this.toolTip1.InitialDelay = 100;
            this.toolTip1.ReshowDelay = 20;
            this.toolTip1.Popup += new System.Windows.Forms.PopupEventHandler(this.toolTip1_Popup);
            // 
            // AudioMix
            // 
            this.AudioMix.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.AudioMix.Cursor = System.Windows.Forms.Cursors.Hand;
            this.AudioMix.FlatAppearance.BorderSize = 0;
            this.AudioMix.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.AudioMix.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.AudioMix.Location = new System.Drawing.Point(127, 145);
            this.AudioMix.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.AudioMix.Name = "AudioMix";
            this.AudioMix.Size = new System.Drawing.Size(89, 29);
            this.AudioMix.TabIndex = 4;
            this.AudioMix.Text = "开始混音";
            this.toolTip1.SetToolTip(this.AudioMix, "将列表中的音频进行混音");
            this.AudioMix.UseVisualStyleBackColor = false;
            this.AudioMix.Click += new System.EventHandler(this.AudioMix_Click);
            // 
            // AddAudioFile
            // 
            this.AddAudioFile.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.AddAudioFile.Cursor = System.Windows.Forms.Cursors.Hand;
            this.AddAudioFile.FlatAppearance.BorderSize = 0;
            this.AddAudioFile.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.AddAudioFile.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.AddAudioFile.Location = new System.Drawing.Point(29, 145);
            this.AddAudioFile.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.AddAudioFile.Name = "AddAudioFile";
            this.AddAudioFile.Size = new System.Drawing.Size(89, 29);
            this.AddAudioFile.TabIndex = 2;
            this.AddAudioFile.Text = "添加音频";
            this.toolTip1.SetToolTip(this.AddAudioFile, "添加音频");
            this.AddAudioFile.UseVisualStyleBackColor = false;
            this.AddAudioFile.Click += new System.EventHandler(this.AddAudioFile_Click);
            // 
            // AudioFilesBox
            // 
            this.AudioFilesBox.AllowDrop = true;
            this.AudioFilesBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(252)))), ((int)(((byte)(252)))), ((int)(((byte)(252)))));
            this.AudioFilesBox.ContextMenuStrip = this.contextMenuStrip1;
            this.AudioFilesBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.AudioFilesBox.FormattingEnabled = true;
            this.AudioFilesBox.HorizontalScrollbar = true;
            this.AudioFilesBox.ItemHeight = 17;
            this.AudioFilesBox.Location = new System.Drawing.Point(10, 33);
            this.AudioFilesBox.Name = "AudioFilesBox";
            this.AudioFilesBox.Size = new System.Drawing.Size(226, 106);
            this.AudioFilesBox.TabIndex = 3;
            this.AudioFilesBox.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.AudioFilesBox_MouseDoubleClick);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.DeleteSelection,
            this.DeleteAll});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(125, 48);
            // 
            // DeleteSelection
            // 
            this.DeleteSelection.Name = "DeleteSelection";
            this.DeleteSelection.Size = new System.Drawing.Size(124, 22);
            this.DeleteSelection.Text = "删除所选";
            this.DeleteSelection.Click += new System.EventHandler(this.DeleteSelection_Click);
            // 
            // DeleteAll
            // 
            this.DeleteAll.Name = "DeleteAll";
            this.DeleteAll.Size = new System.Drawing.Size(124, 22);
            this.DeleteAll.Text = "删除全部";
            this.DeleteAll.Click += new System.EventHandler(this.DeleteAll_Click);
            // 
            // Audio_Mix
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(248)))), ((int)(((byte)(248)))));
            this.ClientSize = new System.Drawing.Size(248, 283);
            this.Controls.Add(this.AudioMix);
            this.Controls.Add(this.AddAudioFile);
            this.Controls.Add(this.AudioFilesBox);
            this.Controls.Add(this.Total_Label);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.Current_Label);
            this.Controls.Add(this.StopButton);
            this.Controls.Add(this.PlayButton);
            this.Controls.Add(this.trackBar1);
            this.Controls.Add(this.TitleLabel);
            this.Controls.Add(this.CloseButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Audio_Mix";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "音频混合";
            this.TopMost = true;
            ((System.ComponentModel.ISupportInitialize)(this.trackBar1)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label TitleLabel;
        private System.Windows.Forms.Button CloseButton;
        private System.Windows.Forms.Label Total_Label;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label Current_Label;
        private System.Windows.Forms.Button StopButton;
        private System.Windows.Forms.Button PlayButton;
        private System.Windows.Forms.TrackBar trackBar1;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.ListBox AudioFilesBox;
        private System.Windows.Forms.Button AudioMix;
        private System.Windows.Forms.Button AddAudioFile;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem DeleteSelection;
        private System.Windows.Forms.ToolStripMenuItem DeleteAll;
    }
}