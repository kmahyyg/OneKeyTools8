using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NAudio.CoreAudioApi;
using NAudio.Wave;
using NAudio.Wave.SampleProviders;
using System.IO;

namespace OneKeyTools
{
    public partial class Audio_Mix : Form
    {
        public Audio_Mix()
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

        private void CloseButton_Click(object sender, EventArgs e)
        {
            if (waveout != null)
            {
                waveout.Stop();
                waveout.PlaybackStopped += OnPlaybackStopped;
            }
            Audio_Mix.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button288.Enabled = true;
        }

        string inputname = "";
        private WaveOut waveout = null;
        private AudioFileReader afr = null;
        int allseconds = 0;

        private void PlayButton_Click(object sender, EventArgs e)
        {
            if (waveout != null && waveout.PlaybackState == PlaybackState.Playing)
            {
                waveout.Pause();
                PlayButton.Text = "已暂停";
                timer1.Stop();
            }
            else if (waveout != null && waveout.PlaybackState == PlaybackState.Paused)
            {
                waveout.Play();
                PlayButton.Text = "播放";
                timer1.Start();
            }
            else
            {
                if (AudioFilesBox.Items.Count != 0)
                {
                    if (inputname == "")
                    {
                        inputname = AudioFilesBox.SelectedItem.ToString();
                    }
                    TitleLabel.Text = Path.GetFileName(inputname);
                    afr = new AudioFileReader(inputname);
                    waveout = new WaveOut();  
                    waveout.Init(afr);
                    waveout.Play();
                    allseconds = afr.TotalTime.Hours * 3600 + afr.TotalTime.Minutes * 60 + afr.TotalTime.Seconds;
                    PlayButton.Text = "播放";
                    timer1.Start();
                    trackBar1.Enabled = true;
                    StopButton.Enabled = true;
                }
                else
                {
                    MessageBox.Show("请先添加音频");
                }
            }
        }

        private void StopButton_Click(object sender, EventArgs e)
        {
            if (waveout != null)
            {
                trackBar1.Value = 0;
                timer1.Stop();
                waveout.Stop();
                waveout.PlaybackStopped += OnPlaybackStopped;
                TitleLabel.Text = "音频混合";
                inputname = "";
            }
        }

        private void OnPlaybackStopped(object sender, StoppedEventArgs e)
        {
            if (afr != null)
            {
                afr.Dispose();
                afr = null;
            }
            if (waveout != null)
            {
                waveout.Dispose();
                waveout = null;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (afr != null)
            {
                if (afr.CurrentTime.ToString().Contains(char.Parse(".")))
                {
                    Current_Label.Text = afr.CurrentTime.ToString().Substring(0, afr.CurrentTime.ToString().LastIndexOf("."));
                }
                else
                {
                    Current_Label.Text = afr.CurrentTime.ToString();
                }
                Total_Label.Text = afr.TotalTime.ToString().Substring(0, afr.TotalTime.ToString().LastIndexOf("."));
                trackBar1.Value = (int)((double)afr.Position / afr.Length * 100);
            }
        }

        int lastvalue = 0;
        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            if (afr != null && trackBar1.Value != lastvalue)
            {
                afr.Skip((int)((trackBar1.Value - lastvalue) * 0.01f * allseconds));
            }
        }

        private void trackBar1_ValueChanged(object sender, EventArgs e)
        {
            if (trackBar1.Value < 100 && trackBar1.Value >= 0)
            {
                lastvalue = trackBar1.Value;
            }
            else
            {
                trackBar1.Value = 0;
                Current_Label.Text = "00:00:00";
                if (waveout != null)
                {
                    waveout.Resume();
                    timer1.Stop();
                    PlayButton.Text = "已暂停";
                }
                else
                {
                    PlayButton.Text = "播放";
                }
            }
        }

        private void AddAudioFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.Filter = "MP3音频(*.mp3)|*.mp3|WAV音频(*.wav)|*.wav";
            of.Multiselect = true;

            if (of.ShowDialog() == DialogResult.OK)
            {
                for (int i = 0; i < of.FileNames.Count(); i++)
                {
                    if (AudioFilesBox.SelectedItems.Count == 0)
                    {
                        AudioFilesBox.Items.Add(of.FileNames[i]);
                        AudioFilesBox.SetSelected(AudioFilesBox.Items.Count - 1, true);
                    }
                    else if (AudioFilesBox.SelectedItems.Count == 1)
                    {
                        int n = AudioFilesBox.SelectedIndex;
                        AudioFilesBox.Items.Insert(n + 1, of.FileNames[i]);
                        AudioFilesBox.SelectedItems.Clear();
                        AudioFilesBox.SetSelected(n + 1, true);
                    }
                }
            }
        }

        private void AudioMix_Click(object sender, EventArgs e)
        {
            if (AudioFilesBox.Items.Count < 2)
            {
                MessageBox.Show("请先插入至少2个音频");
            }
            else
            {
                List<string> Files = new List<string>();
                for (int i = 0; i < AudioFilesBox.Items.Count; i++)
                {
                    Files.Add(AudioFilesBox.Items[i].ToString());
                }

                SaveFileDialog sf = new SaveFileDialog();
                sf.Filter = "MP3音频(*.mp3)|*.mp3";
                sf.AddExtension = true;
                sf.Title = "音频另存为";
                if (sf.ShowDialog() == DialogResult.OK)
                {
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    try
                    {
                        MixAudio(Files.ToArray(), sf.FileName);
                        StopButton_Click(null, null);
                        inputname = sf.FileName;
                        TitleLabel.Text = Path.GetFileName(inputname);
                        afr = new AudioFileReader(inputname);
                        waveout = new WaveOut();
                        waveout.Init(afr);
                        waveout.Play();
                        allseconds = afr.TotalTime.Hours * 3600 + afr.TotalTime.Minutes * 60 + afr.TotalTime.Seconds;
                        PlayButton.Text = "播放";
                        timer1.Start();
                        trackBar1.Enabled = true;
                        StopButton.Enabled = true;
                    }
                    catch
                    {
                        File.Delete(sf.FileName);
                        MessageBox.Show("很遗憾，转换失败");
                    }
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                }
            }
        }

        private void MixAudio(string[] SourceAudios, string outAudio)
        {
            MixingSampleProvider mixer = new MixingSampleProvider(WaveFormat.CreateIeeeFloatWaveFormat(44100, 2));
            NAudio.MediaFoundation.MediaFoundationApi.Startup();
            for (int i = 0; i < SourceAudios.Count(); i++)
            {
                var sourcefile = new AudioFileReader(SourceAudios[i]);
                var mfr = new MediaFoundationResampler(sourcefile, WaveFormat.CreateIeeeFloatWaveFormat(44100, 2));
                mixer.AddMixerInput(mfr);
            }
            var converted16Bit = new SampleToWaveProvider16((ISampleProvider)mixer);
            using (var resampled = new MediaFoundationResampler(converted16Bit, new WaveFormat(44100, 2)))
            {
                MediaFoundationEncoder.EncodeToMp3(resampled, outAudio, 192000);
            }
            NAudio.MediaFoundation.MediaFoundationApi.Shutdown();
        }

        private void DeleteSelection_Click(object sender, EventArgs e)
        {
            if (AudioFilesBox.SelectedItems.Count != 0)
            {
                for (int i = AudioFilesBox.SelectedItems.Count - 1; i >= 0; i--)
                {
                    AudioFilesBox.Items.Remove(AudioFilesBox.SelectedItems[i]);
                }
            }
        }

        private void DeleteAll_Click(object sender, EventArgs e)
        {
            if (AudioFilesBox.Items.Count != 0)
            {
                for (int i = AudioFilesBox.Items.Count - 1; i >= 0; i--)
                {
                    AudioFilesBox.Items.RemoveAt(i);
                }
            }
        }

        private void TitleLabel_Click(object sender, EventArgs e)
        {
            if (inputname != "")
            {
                System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo("Explorer.exe");
                psi.Arguments = "/e,/select," + inputname;
                System.Diagnostics.Process.Start(psi);
            }
        }

        private void toolTip1_Popup(object sender, PopupEventArgs e)
        {
            toolTip1.SetToolTip(TitleLabel, TitleLabel.Text);
        }

        private void AudioFilesBox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (AudioFilesBox.Items.Count != 0 && AudioFilesBox.SelectedItems.Count != 0)
                {
                    inputname = AudioFilesBox.SelectedItem.ToString();
                    PlayButton_Click(null, null);
                }
            }
        }
        
    }
}
