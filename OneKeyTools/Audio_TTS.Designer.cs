namespace OneKeyTools
{
    partial class Audio_TTS
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Audio_TTS));
            this.TitleLabel = new System.Windows.Forms.Label();
            this.CloseButton = new System.Windows.Forms.Button();
            this.SpeakButton = new System.Windows.Forms.Button();
            this.EngineLabel = new System.Windows.Forms.Label();
            this.EngineBox = new System.Windows.Forms.ComboBox();
            this.SpeedLabel = new System.Windows.Forms.Label();
            this.VolumnLabel = new System.Windows.Forms.Label();
            this.trackBar1 = new System.Windows.Forms.TrackBar();
            this.trackBar2 = new System.Windows.Forms.TrackBar();
            this.OutputAllButton = new System.Windows.Forms.Button();
            this.OutputSingleButton = new System.Windows.Forms.Button();
            this.IsToPPT = new System.Windows.Forms.CheckBox();
            this.PauseButton = new System.Windows.Forms.Button();
            this.StopButton = new System.Windows.Forms.Button();
            this.OutputFolder = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.OutTypeBox = new System.Windows.Forms.ComboBox();
            this.ReadNotes = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.trackBar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.trackBar2)).BeginInit();
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
            this.TitleLabel.Text = "朗读工具";
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
            // SpeakButton
            // 
            this.SpeakButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(67)))), ((int)(((byte)(67)))));
            this.SpeakButton.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("SpeakButton.BackgroundImage")));
            this.SpeakButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.SpeakButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.SpeakButton.FlatAppearance.BorderSize = 0;
            this.SpeakButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.SpeakButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.SpeakButton.ForeColor = System.Drawing.Color.White;
            this.SpeakButton.Location = new System.Drawing.Point(76, 166);
            this.SpeakButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.SpeakButton.Name = "SpeakButton";
            this.SpeakButton.Size = new System.Drawing.Size(74, 34);
            this.SpeakButton.TabIndex = 5;
            this.SpeakButton.Tag = "";
            this.toolTip1.SetToolTip(this.SpeakButton, "语音播放选中带文本的形状或页面");
            this.SpeakButton.UseVisualStyleBackColor = false;
            this.SpeakButton.Click += new System.EventHandler(this.SpeakButton_Click);
            // 
            // EngineLabel
            // 
            this.EngineLabel.AutoSize = true;
            this.EngineLabel.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.EngineLabel.Location = new System.Drawing.Point(18, 48);
            this.EngineLabel.Name = "EngineLabel";
            this.EngineLabel.Size = new System.Drawing.Size(32, 17);
            this.EngineLabel.TabIndex = 0;
            this.EngineLabel.Text = "引擎";
            // 
            // EngineBox
            // 
            this.EngineBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.EngineBox.Font = new System.Drawing.Font("微软雅黑", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.EngineBox.FormattingEnabled = true;
            this.EngineBox.Location = new System.Drawing.Point(59, 44);
            this.EngineBox.Name = "EngineBox";
            this.EngineBox.Size = new System.Drawing.Size(149, 24);
            this.EngineBox.TabIndex = 2;
            this.toolTip1.SetToolTip(this.EngineBox, "选择电脑中已安装的语音引擎");
            // 
            // SpeedLabel
            // 
            this.SpeedLabel.AutoSize = true;
            this.SpeedLabel.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.SpeedLabel.Location = new System.Drawing.Point(18, 85);
            this.SpeedLabel.Name = "SpeedLabel";
            this.SpeedLabel.Size = new System.Drawing.Size(32, 17);
            this.SpeedLabel.TabIndex = 0;
            this.SpeedLabel.Text = "语速";
            // 
            // VolumnLabel
            // 
            this.VolumnLabel.AutoSize = true;
            this.VolumnLabel.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.VolumnLabel.Location = new System.Drawing.Point(18, 125);
            this.VolumnLabel.Name = "VolumnLabel";
            this.VolumnLabel.Size = new System.Drawing.Size(32, 17);
            this.VolumnLabel.TabIndex = 0;
            this.VolumnLabel.Text = "音量";
            // 
            // trackBar1
            // 
            this.trackBar1.Location = new System.Drawing.Point(54, 79);
            this.trackBar1.Minimum = -10;
            this.trackBar1.Name = "trackBar1";
            this.trackBar1.Size = new System.Drawing.Size(156, 45);
            this.trackBar1.TabIndex = 3;
            this.trackBar1.TickFrequency = 2;
            this.trackBar1.TickStyle = System.Windows.Forms.TickStyle.TopLeft;
            this.toolTip1.SetToolTip(this.trackBar1, "调整语音的速度");
            // 
            // trackBar2
            // 
            this.trackBar2.Location = new System.Drawing.Point(54, 118);
            this.trackBar2.Maximum = 100;
            this.trackBar2.Name = "trackBar2";
            this.trackBar2.Size = new System.Drawing.Size(156, 45);
            this.trackBar2.SmallChange = 5;
            this.trackBar2.TabIndex = 4;
            this.trackBar2.TickFrequency = 10;
            this.trackBar2.TickStyle = System.Windows.Forms.TickStyle.TopLeft;
            this.toolTip1.SetToolTip(this.trackBar2, "调整语音的音量");
            this.trackBar2.Value = 100;
            // 
            // OutputAllButton
            // 
            this.OutputAllButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.OutputAllButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.OutputAllButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.OutputAllButton.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.OutputAllButton.FlatAppearance.BorderSize = 0;
            this.OutputAllButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White;
            this.OutputAllButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.OutputAllButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.OutputAllButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.OutputAllButton.ForeColor = System.Drawing.Color.Black;
            this.OutputAllButton.Location = new System.Drawing.Point(21, 270);
            this.OutputAllButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.OutputAllButton.Name = "OutputAllButton";
            this.OutputAllButton.Size = new System.Drawing.Size(88, 34);
            this.OutputAllButton.TabIndex = 10;
            this.OutputAllButton.Text = "合并导出";
            this.toolTip1.SetToolTip(this.OutputAllButton, "将语音合并为一个音频");
            this.OutputAllButton.UseVisualStyleBackColor = false;
            this.OutputAllButton.Click += new System.EventHandler(this.OutputAllButton_Click);
            // 
            // OutputSingleButton
            // 
            this.OutputSingleButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.OutputSingleButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.OutputSingleButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.OutputSingleButton.FlatAppearance.BorderColor = System.Drawing.Color.DimGray;
            this.OutputSingleButton.FlatAppearance.BorderSize = 0;
            this.OutputSingleButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White;
            this.OutputSingleButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.White;
            this.OutputSingleButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.OutputSingleButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.OutputSingleButton.ForeColor = System.Drawing.Color.Black;
            this.OutputSingleButton.Location = new System.Drawing.Point(117, 270);
            this.OutputSingleButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.OutputSingleButton.Name = "OutputSingleButton";
            this.OutputSingleButton.Size = new System.Drawing.Size(88, 34);
            this.OutputSingleButton.TabIndex = 11;
            this.OutputSingleButton.Text = "独立导出";
            this.toolTip1.SetToolTip(this.OutputSingleButton, "将语音分别独立导出为音频");
            this.OutputSingleButton.UseVisualStyleBackColor = false;
            this.OutputSingleButton.Click += new System.EventHandler(this.OutputSingleButton_Click);
            // 
            // IsToPPT
            // 
            this.IsToPPT.AutoSize = true;
            this.IsToPPT.BackColor = System.Drawing.Color.Transparent;
            this.IsToPPT.Checked = true;
            this.IsToPPT.CheckState = System.Windows.Forms.CheckState.Checked;
            this.IsToPPT.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.IsToPPT.Location = new System.Drawing.Point(40, 241);
            this.IsToPPT.Name = "IsToPPT";
            this.IsToPPT.Size = new System.Drawing.Size(72, 21);
            this.IsToPPT.TabIndex = 8;
            this.IsToPPT.Text = "导入PPT";
            this.toolTip1.SetToolTip(this.IsToPPT, "勾选后导出的音频导入到PPT");
            this.IsToPPT.UseVisualStyleBackColor = false;
            // 
            // PauseButton
            // 
            this.PauseButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.PauseButton.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("PauseButton.BackgroundImage")));
            this.PauseButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.PauseButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PauseButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.PauseButton.FlatAppearance.BorderSize = 0;
            this.PauseButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.PauseButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.PauseButton.ForeColor = System.Drawing.Color.Black;
            this.PauseButton.Location = new System.Drawing.Point(21, 166);
            this.PauseButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.PauseButton.Name = "PauseButton";
            this.PauseButton.Size = new System.Drawing.Size(49, 34);
            this.PauseButton.TabIndex = 6;
            this.toolTip1.SetToolTip(this.PauseButton, "暂停正在播放的语音");
            this.PauseButton.UseVisualStyleBackColor = false;
            this.PauseButton.Click += new System.EventHandler(this.PauseButton_Click);
            // 
            // StopButton
            // 
            this.StopButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.StopButton.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("StopButton.BackgroundImage")));
            this.StopButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.StopButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.StopButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.StopButton.FlatAppearance.BorderSize = 0;
            this.StopButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.StopButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.StopButton.ForeColor = System.Drawing.Color.Black;
            this.StopButton.Location = new System.Drawing.Point(156, 166);
            this.StopButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.StopButton.Name = "StopButton";
            this.StopButton.Size = new System.Drawing.Size(49, 34);
            this.StopButton.TabIndex = 7;
            this.toolTip1.SetToolTip(this.StopButton, "停止正在播放的语音");
            this.StopButton.UseVisualStyleBackColor = false;
            this.StopButton.Click += new System.EventHandler(this.StopButton_Click);
            // 
            // OutputFolder
            // 
            this.OutputFolder.AutoSize = true;
            this.OutputFolder.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.OutputFolder.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.OutputFolder.Location = new System.Drawing.Point(79, 312);
            this.OutputFolder.Name = "OutputFolder";
            this.OutputFolder.Size = new System.Drawing.Size(68, 17);
            this.OutputFolder.TabIndex = 12;
            this.OutputFolder.Text = "语音文件夹";
            this.toolTip1.SetToolTip(this.OutputFolder, "打开语音文件夹");
            this.OutputFolder.Click += new System.EventHandler(this.OutputFolder_Click);
            // 
            // toolTip1
            // 
            this.toolTip1.AutomaticDelay = 100;
            this.toolTip1.AutoPopDelay = 3000;
            this.toolTip1.InitialDelay = 100;
            this.toolTip1.ReshowDelay = 20;
            // 
            // OutTypeBox
            // 
            this.OutTypeBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.OutTypeBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.OutTypeBox.FormattingEnabled = true;
            this.OutTypeBox.Items.AddRange(new object[] {
            "mp3",
            "wav"});
            this.OutTypeBox.Location = new System.Drawing.Point(118, 239);
            this.OutTypeBox.Name = "OutTypeBox";
            this.OutTypeBox.Size = new System.Drawing.Size(77, 25);
            this.OutTypeBox.TabIndex = 9;
            this.OutTypeBox.SelectedIndexChanged += new System.EventHandler(this.OutTypeBox_SelectedIndexChanged);
            // 
            // ReadNotes
            // 
            this.ReadNotes.AutoSize = true;
            this.ReadNotes.BackColor = System.Drawing.Color.Transparent;
            this.ReadNotes.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ReadNotes.Location = new System.Drawing.Point(77, 210);
            this.ReadNotes.Name = "ReadNotes";
            this.ReadNotes.Size = new System.Drawing.Size(75, 21);
            this.ReadNotes.TabIndex = 8;
            this.ReadNotes.Text = "朗读备注";
            this.ReadNotes.UseVisualStyleBackColor = false;
            // 
            // Audio_TTS
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(248)))), ((int)(((byte)(248)))));
            this.ClientSize = new System.Drawing.Size(226, 343);
            this.Controls.Add(this.StopButton);
            this.Controls.Add(this.PauseButton);
            this.Controls.Add(this.SpeakButton);
            this.Controls.Add(this.OutTypeBox);
            this.Controls.Add(this.ReadNotes);
            this.Controls.Add(this.IsToPPT);
            this.Controls.Add(this.trackBar2);
            this.Controls.Add(this.trackBar1);
            this.Controls.Add(this.EngineBox);
            this.Controls.Add(this.VolumnLabel);
            this.Controls.Add(this.SpeedLabel);
            this.Controls.Add(this.OutputFolder);
            this.Controls.Add(this.EngineLabel);
            this.Controls.Add(this.OutputSingleButton);
            this.Controls.Add(this.OutputAllButton);
            this.Controls.Add(this.TitleLabel);
            this.Controls.Add(this.CloseButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Audio_TTS";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "朗读工具";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.Audio_TTS_Load);
            ((System.ComponentModel.ISupportInitialize)(this.trackBar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.trackBar2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label TitleLabel;
        private System.Windows.Forms.Button CloseButton;
        private System.Windows.Forms.Button SpeakButton;
        private System.Windows.Forms.Label EngineLabel;
        private System.Windows.Forms.ComboBox EngineBox;
        private System.Windows.Forms.Label SpeedLabel;
        private System.Windows.Forms.Label VolumnLabel;
        private System.Windows.Forms.TrackBar trackBar1;
        private System.Windows.Forms.TrackBar trackBar2;
        private System.Windows.Forms.Button OutputAllButton;
        private System.Windows.Forms.Button OutputSingleButton;
        private System.Windows.Forms.CheckBox IsToPPT;
        private System.Windows.Forms.Button PauseButton;
        private System.Windows.Forms.Button StopButton;
        private System.Windows.Forms.Label OutputFolder;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.ComboBox OutTypeBox;
        private System.Windows.Forms.CheckBox ReadNotes;
    }
}