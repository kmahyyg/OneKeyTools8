using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Speech.Recognition;
using System.Globalization;
using NAudio;
using NAudio.Wave;
using System.IO;

namespace OneKeyTools
{
    public partial class Audio_Record : Form
    {
        public Audio_Record()
        {
            InitializeComponent();

        }

        PowerPoint.Application app = Globals.ThisAddIn.Application;

        private bool m_isMouseDown = false;
        private Point m_mousePos = new Point();
        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            m_mousePos = Cursor.Position;
            m_isMouseDown = true;
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            m_isMouseDown = false;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (m_isMouseDown)
            {
                Point tempPos = Cursor.Position;
                this.Location = new Point(Location.X + (tempPos.X - m_mousePos.X), Location.Y + (tempPos.Y - m_mousePos.Y));
                m_mousePos = Cursor.Position;
            }
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            Audio_Record.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button286.Enabled = true;
        }

        private void Audio_Record_Load(object sender, EventArgs e)
        {
            pname = app.ActivePresentation.Name;
            if (pname.Contains(".pptx"))
            {
                pname = pname.Replace(".pptx", "");
            }
            if (pname.Contains(".ppt"))
            {
                pname = pname.Replace(".ppt", "");
            }
            cPath = app.ActivePresentation.Path + @"\" + pname + @" 的录音\";
        }

        string pname = "";
        string cPath = "";
        string wavename = "";

        private void BeginRecognition_Click(object sender, EventArgs e)
        {
            if (TitleLabel.Text == "录音工具")
            {
                if (!Directory.Exists(cPath))
                {
                    Directory.CreateDirectory(cPath);
                }
                System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(cPath);
                wavename = cPath + "录音文件_" + dir.GetFiles().Length + ".wav";
                waveIn = new WaveInEvent();
                waveIn.WaveFormat = new WaveFormat(44100, 1);
                waveIn.DataAvailable += new EventHandler<WaveInEventArgs>(waveIn_DataAvailable);
                waveIn.RecordingStopped += new EventHandler<StoppedEventArgs>(OnRecordingStopped);
                writer = new WaveFileWriter(wavename, waveIn.WaveFormat);
                waveIn.StartRecording();
                TitleLabel.Text = "录音中";
            }
            else
            {
                if (waveIn != null)
                {
                    waveIn.StopRecording();
                    RecordAudiosBox.Items.Add(wavename);
                    TitleLabel.Text = "录音工具";
                }
            } 
        }

        //录音

        private static WaveInEvent waveIn;
        private static WaveFileWriter writer;

        private void waveIn_DataAvailable(object sender, WaveInEventArgs e)
        {
            if (waveIn != null)
            {
                writer.Write(e.Buffer, 0, e.BytesRecorded);
                writer.Flush();
            }
        }

        private void OnRecordingStopped(object sender, StoppedEventArgs e)
        {
            if (waveIn != null)
            {
                waveIn.Dispose();
                waveIn = null;
            }
            if (writer != null)
            {
                writer.Close();
                writer = null;
            }
            if (e.Exception != null)
            {
                MessageBox.Show(String.Format("出现问题 {0}", e.Exception.Message));
            }
            TitleLabel.Text = "录音工具";
        }

        private void TitleLabel_Click(object sender, EventArgs e)
        {
            if (cPath != "")
            {
                System.Diagnostics.Process.Start("Explorer.exe", cPath);
            }
            else
            {
                MessageBox.Show("不存在语音文件夹，请先录制语音");
            }
        }

        private void 删除所选ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (RecordAudiosBox.SelectedItems.Count != 0)
            {
                for (int i = RecordAudiosBox.SelectedItems.Count - 1; i >= 0; i--)
                {
                    try
                    {
                        File.Delete(RecordAudiosBox.SelectedItems[i].ToString());
                    }
                    catch
                    { } 
                    RecordAudiosBox.Items.Remove(RecordAudiosBox.SelectedItems[i]);

                }
            }
        }

        private void 删除所有ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (RecordAudiosBox.Items.Count != 0)
            {
                for (int i = RecordAudiosBox.Items.Count - 1; i >= 0; i--)
                {
                    try
                    {
                        File.Delete(RecordAudiosBox.Items[i].ToString());
                    }
                    catch
                    { } 
                    RecordAudiosBox.Items.RemoveAt(i);

                }
            }
        }

        private void RecordAudiosBox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (RecordAudiosBox.Items.Count != 0 && RecordAudiosBox.SelectedItems.Count != 0)
                {
                    System.Diagnostics.Process.Start("Explorer.exe", RecordAudiosBox.SelectedItem.ToString());
                }
            }
        }

    }
}
