namespace OneKeyTools
{
    partial class Audio_Split
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
            this.OpenButton = new System.Windows.Forms.Button();
            this.PlayButton = new System.Windows.Forms.Button();
            this.StopButton = new System.Windows.Forms.Button();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.Current_Label = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.Total_Label = new System.Windows.Forms.Label();
            this.trackBar1 = new System.Windows.Forms.TrackBar();
            this.TimeLabelsBox = new System.Windows.Forms.ListBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.DeleteLabelsMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.DeleteAllLabelsMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.重新排序ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.AddTimeLabel = new System.Windows.Forms.Button();
            this.SplitAudioButtonAll = new System.Windows.Forms.Button();
            this.CombinButton = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.button2 = new System.Windows.Forms.Button();
            this.toolTip2 = new System.Windows.Forms.ToolTip(this.components);
            this.SplitAudioButtonSelected = new System.Windows.Forms.Button();
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
            this.TitleLabel.Text = "音频拆合";
            this.TitleLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.toolTip1.SetToolTip(this.TitleLabel, "音频拆合");
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
            this.CloseButton.TabIndex = 10;
            this.CloseButton.Text = "关闭";
            this.toolTip2.SetToolTip(this.CloseButton, "关闭界面");
            this.CloseButton.UseVisualStyleBackColor = false;
            this.CloseButton.Click += new System.EventHandler(this.CloseButton_Click);
            // 
            // OpenButton
            // 
            this.OpenButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.OpenButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.OpenButton.FlatAppearance.BorderSize = 0;
            this.OpenButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.OpenButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.OpenButton.Location = new System.Drawing.Point(22, 97);
            this.OpenButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.OpenButton.Name = "OpenButton";
            this.OpenButton.Size = new System.Drawing.Size(64, 32);
            this.OpenButton.TabIndex = 1;
            this.OpenButton.Text = "加载";
            this.toolTip2.SetToolTip(this.OpenButton, "加载音频");
            this.OpenButton.UseVisualStyleBackColor = false;
            this.OpenButton.Click += new System.EventHandler(this.OpenButton_Click);
            // 
            // PlayButton
            // 
            this.PlayButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(67)))), ((int)(((byte)(67)))));
            this.PlayButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.PlayButton.Enabled = false;
            this.PlayButton.FlatAppearance.BorderSize = 0;
            this.PlayButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.PlayButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.PlayButton.ForeColor = System.Drawing.Color.White;
            this.PlayButton.Location = new System.Drawing.Point(93, 97);
            this.PlayButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.PlayButton.Name = "PlayButton";
            this.PlayButton.Size = new System.Drawing.Size(64, 32);
            this.PlayButton.TabIndex = 2;
            this.PlayButton.Text = "播放";
            this.toolTip2.SetToolTip(this.PlayButton, "播放/暂停音频");
            this.PlayButton.UseVisualStyleBackColor = false;
            this.PlayButton.Click += new System.EventHandler(this.PlayButton_Click);
            // 
            // StopButton
            // 
            this.StopButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.StopButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.StopButton.Enabled = false;
            this.StopButton.FlatAppearance.BorderSize = 0;
            this.StopButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.StopButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.StopButton.Location = new System.Drawing.Point(164, 97);
            this.StopButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.StopButton.Name = "StopButton";
            this.StopButton.Size = new System.Drawing.Size(64, 32);
            this.StopButton.TabIndex = 4;
            this.StopButton.Text = "停止";
            this.toolTip2.SetToolTip(this.StopButton, "停止播放音频");
            this.StopButton.UseVisualStyleBackColor = false;
            this.StopButton.Click += new System.EventHandler(this.StopButton_Click);
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // Current_Label
            // 
            this.Current_Label.AutoSize = true;
            this.Current_Label.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Current_Label.Location = new System.Drawing.Point(61, 71);
            this.Current_Label.Name = "Current_Label";
            this.Current_Label.Size = new System.Drawing.Size(56, 17);
            this.Current_Label.TabIndex = 0;
            this.Current_Label.Text = "00:00:00";
            this.toolTip2.SetToolTip(this.Current_Label, "当前时间");
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(117, 71);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(13, 17);
            this.label2.TabIndex = 0;
            this.label2.Text = "/";
            // 
            // Total_Label
            // 
            this.Total_Label.AutoSize = true;
            this.Total_Label.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Total_Label.Location = new System.Drawing.Point(130, 71);
            this.Total_Label.Name = "Total_Label";
            this.Total_Label.Size = new System.Drawing.Size(56, 17);
            this.Total_Label.TabIndex = 0;
            this.Total_Label.Text = "00:00:00";
            this.toolTip2.SetToolTip(this.Total_Label, "总时间");
            // 
            // trackBar1
            // 
            this.trackBar1.Cursor = System.Windows.Forms.Cursors.SizeWE;
            this.trackBar1.Enabled = false;
            this.trackBar1.Location = new System.Drawing.Point(17, 44);
            this.trackBar1.Maximum = 100;
            this.trackBar1.Name = "trackBar1";
            this.trackBar1.Size = new System.Drawing.Size(216, 45);
            this.trackBar1.TabIndex = 3;
            this.trackBar1.TickStyle = System.Windows.Forms.TickStyle.None;
            this.trackBar1.Scroll += new System.EventHandler(this.trackBar1_Scroll);
            this.trackBar1.ValueChanged += new System.EventHandler(this.trackBar1_ValueChanged);
            // 
            // TimeLabelsBox
            // 
            this.TimeLabelsBox.AllowDrop = true;
            this.TimeLabelsBox.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(252)))), ((int)(((byte)(252)))), ((int)(((byte)(252)))));
            this.TimeLabelsBox.ContextMenuStrip = this.contextMenuStrip1;
            this.TimeLabelsBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.TimeLabelsBox.FormattingEnabled = true;
            this.TimeLabelsBox.ItemHeight = 17;
            this.TimeLabelsBox.Location = new System.Drawing.Point(22, 142);
            this.TimeLabelsBox.Name = "TimeLabelsBox";
            this.TimeLabelsBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.TimeLabelsBox.Size = new System.Drawing.Size(151, 106);
            this.TimeLabelsBox.TabIndex = 6;
            this.TimeLabelsBox.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.TimeLabelsBox_DrawItem);
            this.TimeLabelsBox.DragDrop += new System.Windows.Forms.DragEventHandler(this.TimeLabelsBox_DragDrop);
            this.TimeLabelsBox.DragOver += new System.Windows.Forms.DragEventHandler(this.TimeLabelsBox_DragOver);
            this.TimeLabelsBox.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.TimeLabelsBox_MouseDoubleClick);
            this.TimeLabelsBox.MouseDown += new System.Windows.Forms.MouseEventHandler(this.TimeLabelsBox_MouseDown);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.DeleteLabelsMenu,
            this.DeleteAllLabelsMenu,
            this.重新排序ToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(125, 70);
            // 
            // DeleteLabelsMenu
            // 
            this.DeleteLabelsMenu.Name = "DeleteLabelsMenu";
            this.DeleteLabelsMenu.Size = new System.Drawing.Size(124, 22);
            this.DeleteLabelsMenu.Text = "删除所选";
            this.DeleteLabelsMenu.Click += new System.EventHandler(this.DeleteLabelsMenu_Click);
            // 
            // DeleteAllLabelsMenu
            // 
            this.DeleteAllLabelsMenu.Name = "DeleteAllLabelsMenu";
            this.DeleteAllLabelsMenu.Size = new System.Drawing.Size(124, 22);
            this.DeleteAllLabelsMenu.Text = "删除所有";
            this.DeleteAllLabelsMenu.Click += new System.EventHandler(this.DeleteAllLabelsMenu_Click);
            // 
            // 重新排序ToolStripMenuItem
            // 
            this.重新排序ToolStripMenuItem.Name = "重新排序ToolStripMenuItem";
            this.重新排序ToolStripMenuItem.Size = new System.Drawing.Size(124, 22);
            this.重新排序ToolStripMenuItem.Text = "重新排序";
            this.重新排序ToolStripMenuItem.Click += new System.EventHandler(this.重新排序ToolStripMenuItem_Click);
            // 
            // AddTimeLabel
            // 
            this.AddTimeLabel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.AddTimeLabel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.AddTimeLabel.FlatAppearance.BorderSize = 0;
            this.AddTimeLabel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.AddTimeLabel.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.AddTimeLabel.Location = new System.Drawing.Point(179, 142);
            this.AddTimeLabel.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.AddTimeLabel.Name = "AddTimeLabel";
            this.AddTimeLabel.Size = new System.Drawing.Size(49, 106);
            this.AddTimeLabel.TabIndex = 5;
            this.AddTimeLabel.Text = "添加\r\n标签";
            this.toolTip2.SetToolTip(this.AddTimeLabel, "添加时间标签");
            this.AddTimeLabel.UseVisualStyleBackColor = false;
            this.AddTimeLabel.Click += new System.EventHandler(this.AddTimeLabel_Click);
            // 
            // SplitAudioButtonAll
            // 
            this.SplitAudioButtonAll.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.SplitAudioButtonAll.Cursor = System.Windows.Forms.Cursors.Hand;
            this.SplitAudioButtonAll.FlatAppearance.BorderSize = 0;
            this.SplitAudioButtonAll.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.SplitAudioButtonAll.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.SplitAudioButtonAll.ForeColor = System.Drawing.Color.Black;
            this.SplitAudioButtonAll.Location = new System.Drawing.Point(128, 257);
            this.SplitAudioButtonAll.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.SplitAudioButtonAll.Name = "SplitAudioButtonAll";
            this.SplitAudioButtonAll.Size = new System.Drawing.Size(100, 30);
            this.SplitAudioButtonAll.TabIndex = 8;
            this.SplitAudioButtonAll.Text = "分割全部";
            this.toolTip2.SetToolTip(this.SplitAudioButtonAll, "根据所有标签分割音频");
            this.SplitAudioButtonAll.UseVisualStyleBackColor = false;
            this.SplitAudioButtonAll.Click += new System.EventHandler(this.SplitAudioButtonAll_Click);
            // 
            // CombinButton
            // 
            this.CombinButton.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.CombinButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.CombinButton.FlatAppearance.BorderSize = 0;
            this.CombinButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.CombinButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CombinButton.Location = new System.Drawing.Point(22, 292);
            this.CombinButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.CombinButton.Name = "CombinButton";
            this.CombinButton.Size = new System.Drawing.Size(206, 32);
            this.CombinButton.TabIndex = 9;
            this.CombinButton.Text = "合并音频";
            this.toolTip2.SetToolTip(this.CombinButton, "从对话框中选中并合并音频");
            this.CombinButton.UseVisualStyleBackColor = false;
            this.CombinButton.Click += new System.EventHandler(this.CombinButton_Click);
            // 
            // toolTip1
            // 
            this.toolTip1.AutomaticDelay = 100;
            this.toolTip1.AutoPopDelay = 2000;
            this.toolTip1.InitialDelay = 100;
            this.toolTip1.ReshowDelay = 20;
            this.toolTip1.Popup += new System.Windows.Forms.PopupEventHandler(this.toolTip1_Popup);
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.button2.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button2.FlatAppearance.BorderSize = 0;
            this.button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button2.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button2.Location = new System.Drawing.Point(56, 27);
            this.button2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(0, 0);
            this.button2.TabIndex = 2;
            this.button2.Text = "打开";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.OpenButton_Click);
            // 
            // toolTip2
            // 
            this.toolTip2.AutomaticDelay = 100;
            this.toolTip2.AutoPopDelay = 2000;
            this.toolTip2.InitialDelay = 100;
            this.toolTip2.ReshowDelay = 20;
            // 
            // SplitAudioButtonSelected
            // 
            this.SplitAudioButtonSelected.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.SplitAudioButtonSelected.Cursor = System.Windows.Forms.Cursors.Hand;
            this.SplitAudioButtonSelected.FlatAppearance.BorderSize = 0;
            this.SplitAudioButtonSelected.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.SplitAudioButtonSelected.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.SplitAudioButtonSelected.ForeColor = System.Drawing.Color.Black;
            this.SplitAudioButtonSelected.Location = new System.Drawing.Point(22, 257);
            this.SplitAudioButtonSelected.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.SplitAudioButtonSelected.Name = "SplitAudioButtonSelected";
            this.SplitAudioButtonSelected.Size = new System.Drawing.Size(100, 30);
            this.SplitAudioButtonSelected.TabIndex = 7;
            this.SplitAudioButtonSelected.Text = "分割所选";
            this.toolTip2.SetToolTip(this.SplitAudioButtonSelected, "分割所选时间标签范围内的音频");
            this.SplitAudioButtonSelected.UseVisualStyleBackColor = false;
            this.SplitAudioButtonSelected.Click += new System.EventHandler(this.SplitAudioButtonSelected_Click);
            // 
            // Audio_Split
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(248)))), ((int)(((byte)(248)))), ((int)(((byte)(248)))));
            this.ClientSize = new System.Drawing.Size(248, 334);
            this.Controls.Add(this.TimeLabelsBox);
            this.Controls.Add(this.Total_Label);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.Current_Label);
            this.Controls.Add(this.CombinButton);
            this.Controls.Add(this.SplitAudioButtonSelected);
            this.Controls.Add(this.SplitAudioButtonAll);
            this.Controls.Add(this.AddTimeLabel);
            this.Controls.Add(this.StopButton);
            this.Controls.Add(this.PlayButton);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.OpenButton);
            this.Controls.Add(this.TitleLabel);
            this.Controls.Add(this.CloseButton);
            this.Controls.Add(this.trackBar1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Audio_Split";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "音频拆合";
            this.TopMost = true;
            ((System.ComponentModel.ISupportInitialize)(this.trackBar1)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label TitleLabel;
        private System.Windows.Forms.Button CloseButton;
        private System.Windows.Forms.Button OpenButton;
        private System.Windows.Forms.Button PlayButton;
        private System.Windows.Forms.Button StopButton;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Label Current_Label;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label Total_Label;
        private System.Windows.Forms.TrackBar trackBar1;
        private System.Windows.Forms.ListBox TimeLabelsBox;
        private System.Windows.Forms.Button AddTimeLabel;
        private System.Windows.Forms.Button SplitAudioButtonAll;
        private System.Windows.Forms.Button CombinButton;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem DeleteLabelsMenu;
        private System.Windows.Forms.ToolStripMenuItem DeleteAllLabelsMenu;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ToolTip toolTip2;
        private System.Windows.Forms.Button SplitAudioButtonSelected;
        private System.Windows.Forms.ToolStripMenuItem 重新排序ToolStripMenuItem;
    }
}