using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OneKeyTools
{
    public partial class Audio_Convert : Form
    {
        public Audio_Convert()
        {
            InitializeComponent();
        }

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

        private static void WavToMp3(string WavFile, string Mp3File)
        {
            NAudio.MediaFoundation.MediaFoundationApi.Startup();
            var mediaType = MediaFoundationEncoder.SelectMediaType(NAudio.MediaFoundation.AudioSubtypes.MFAudioFormat_WMAudioV8, new WaveFormat(16000, 1), 16000);
            if (mediaType != null)
            {
                using (var wavreader = new WaveFileReader(WavFile))
                {
                    MediaFoundationEncoder.EncodeToMp3(wavreader, Mp3File, 48000);
                }
            }
            NAudio.MediaFoundation.MediaFoundationApi.Shutdown();
        }

        private void Mp3ToWav(string mp3file, string wavfile)
        {
            using (Mp3FileReader reader = new Mp3FileReader(mp3file))
            {
                WaveFileWriter.CreateWaveFile(wavfile, reader);
            }
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            Audio_Convert.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button289.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.Multiselect = true;
            of.AddExtension = true;
            of.Filter = "WAV音频(*.wav)|*.wav";
            of.Multiselect = true;
            of.Title = "选择WAV";

            if (of.ShowDialog() == DialogResult.OK)
            {
                string ofname = System.IO.Path.GetFileNameWithoutExtension(of.FileName);
                SaveFileDialog sf = new SaveFileDialog();
                sf.Filter = "MP3音频(*.mp3)|*.mp3";
                sf.AddExtension = true;
                sf.FileName = ofname;
                sf.Title = "音频另存为";
                if (sf.ShowDialog() == DialogResult.OK)
                {
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    int n = 0;
                    string extension = Path.GetExtension(sf.FileName);
                    string sfname = Path.GetFileNameWithoutExtension(sf.FileName);
                    string fpath = Path.GetDirectoryName(sf.FileName);
                    List<string> sfnames = new List<string>();
                    if (sfname == ofname)
                    {
                        for (int i = 0; i < of.FileNames.Count(); i++)
                        {
                            sfnames.Add(fpath + "\\" + Path.GetFileNameWithoutExtension(of.FileNames[i]) + extension);
                        }
                    }
                    else
                    {
                        for (int i = 0; i < of.FileNames.Count(); i++)
                        {
                            sfnames.Add(fpath + "\\" + sfname + "_" + (i + 1) + extension);
                        }
                    }

                    if (of.FileNames.Count() == 1)
                    {
                        try
                        {
                            WavToMp3(of.FileName, sf.FileName);
                        }
                        catch
                        {
                            MessageBox.Show("发生错误，无法转换");
                        }
                    }
                    else
                    {
                        for (int i = 0; i < of.FileNames.Count(); i++)
                        {
                            try
                            {
                                WavToMp3(of.FileNames[i], sfnames[i]);
                                n += 1;
                            }
                            catch
                            {
                                continue;
                            }
                        }
                        MessageBox.Show("转换成功 " + n + " 个音频");
                    }
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    System.Diagnostics.Process.Start("Explorer.exe", fpath);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.Multiselect = true;
            of.AddExtension = true;
            of.Filter = "MP3音频(*.mp3)|*.mp3";
            of.Multiselect = true;
            of.Title = "选择MP3";

            if (of.ShowDialog() == DialogResult.OK)
            {
                string ofname = System.IO.Path.GetFileNameWithoutExtension(of.FileName);
                SaveFileDialog sf = new SaveFileDialog();
                sf.Filter = "WAV音频(*.wav)|*.wav";
                sf.AddExtension = true;
                sf.FileName = ofname;
                sf.Title = "音频另存为";
                if (sf.ShowDialog() == DialogResult.OK)
                {
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    int n = 0;
                    string extension = Path.GetExtension(sf.FileName);
                    string sfname = Path.GetFileNameWithoutExtension(sf.FileName);
                    string fpath = Path.GetDirectoryName(sf.FileName);
                    List<string> sfnames = new List<string>();
                    if (sfname == ofname)
                    {
                        for (int i = 0; i < of.FileNames.Count(); i++)
                        {
                            sfnames.Add(fpath + "\\" + Path.GetFileNameWithoutExtension(of.FileNames[i]) + extension);
                        }
                    }
                    else
                    {
                        for (int i = 0; i < of.FileNames.Count(); i++)
                        {
                            sfnames.Add(fpath + "\\" + sfname + "_" + (i + 1) + extension);
                        }
                    }

                    if (of.FileNames.Count() == 1)
                    {
                        try
                        {
                            Mp3ToWav(of.FileName, sf.FileName);
                            System.Diagnostics.Process.Start("Explorer.exe", sf.FileName);
                        }
                        catch
                        {
                            MessageBox.Show("发生错误，无法转换");
                        }
                    }
                    else
                    {
                        for (int i = 0; i < of.FileNames.Count(); i++)
                        {
                            try
                            {
                                Mp3ToWav(of.FileNames[i], sfnames[i]);
                                n += 1;
                            }
                            catch
                            {
                                continue;
                            }
                        }
                        MessageBox.Show("转换成功 " + n + " 个音频");
                    }
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                }
            }
        }
    }
}
