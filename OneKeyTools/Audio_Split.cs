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
using NAudio.Wave.WaveFormats;
using System.IO;
using System.Threading;

namespace OneKeyTools
{
    public partial class Audio_Split : Form
    {
        public Audio_Split()
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

        private void TimeLabelsBox_MouseDown(object sender, MouseEventArgs e)
        {
            //base.OnMouseDown(e);
            //if (listBox1.Items.Count == 0 || e.Button != MouseButtons.Left || listBox1.SelectedIndex == -1 || e.Clicks == 2)
            //{
            //    return;
            //}
            //else
            //{
            //    int index = listBox1.SelectedIndex;
            //    object item = listBox1.Items[index];
            //    DragDropEffects dde = DoDragDrop(item, DragDropEffects.Move);
            //}
        }

        private void TimeLabelsBox_DragOver(object sender, DragEventArgs e)
        {
            //base.OnDragOver(e);
            //e.Effect = DragDropEffects.Move;
        }

        private void TimeLabelsBox_DragDrop(object sender, DragEventArgs e)
        {
            //base.OnDragDrop(e);
            //object item = listBox1.SelectedItem;
            //int index = listBox1.IndexFromPoint(this.PointToClient(new Point(e.X, e.Y)));
            //listBox1.Items.Remove(item);
            //if (index < 0)
            //{
            //    listBox1.Items.Add(item);
            //}
            //else
            //{
            //    listBox1.Items.Insert(index, item);
            //}
        }

        private void TimeLabelsBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            //if (listBox1.Items.Count < 0) return;
            //if (e.Index < 0) return;
            //bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            //if (selected)
            //    e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(255,255,255,255)), e.Bounds);
            //else if (e.Index % 2 != 0)
            //    e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(255, 252, 252, 252)), e.Bounds);
            //else
            //    e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(255, 248, 248, 248)), e.Bounds);

            ////if (selected)
            ////{
            ////    e.Graphics.DrawString(this.GetItemText(e.Index), e.Font,selectFontBursh, e.Bounds);
            ////}
            ////else
            ////{
            ////    e.Graphics.DrawString(this.GetItemText(e.Index), e.Font,normalFontBursh, e.Bounds);
            ////}
            //e.DrawFocusRectangle();
            //listBox1_DrawItem(sender,e);
        }

        string inputname = "";
        private WaveOut waveout = null;
        private AudioFileReader afr = null;
        int allseconds = 0;
        
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

        private void CloseButton_Click(object sender, EventArgs e)
        {
            if (waveout != null)
            {
                waveout.Stop();
                waveout.PlaybackStopped += OnPlaybackStopped;
            }
            Audio_Split.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button284.Enabled = true;
        }

        private void OpenButton_Click(object sender, EventArgs e)
        {
            //本功能使用开源免费库NAudio http://naudio.codeplex.com/
            OpenFileDialog of = new OpenFileDialog();
            of.Multiselect = true;
            of.Filter = "MP3音频(*.mp3)|*.mp3|WAV音频(*.wav)|*.wav";

            if (of.ShowDialog() == DialogResult.OK)
            {
                inputname = of.FileName;
                if (waveout != null)
                {
                    waveout.Stop();
                    waveout.PlaybackStopped += OnPlaybackStopped;
                }
                trackBar1.Enabled = true;
                trackBar1.Value = 0;
                PlayButton.Enabled = true;
                StopButton.Enabled = true;
                TitleLabel.Text = Path.GetFileName(inputname);
                TimeLabelsBox.Items.Clear();
                this.PlayButton_Click(null,null);
            }
        }

        private void PlayButton_Click(object sender, EventArgs e)
        {
            if (waveout != null && waveout.PlaybackState == PlaybackState.Playing)
            {
                waveout.Pause();
                PlayButton.Text = "已暂停";
                timer1.Stop();
            }
            else if(waveout != null && waveout.PlaybackState == PlaybackState.Paused)
            {
                waveout.Play();
                PlayButton.Text = "播放";
                timer1.Start();
            }
            else
            {
                string extension = System.IO.Path.GetExtension(inputname);
                afr = new AudioFileReader(inputname);
                waveout = new WaveOut();
                allseconds = afr.TotalTime.Hours * 3600 + afr.TotalTime.Minutes * 60 + afr.TotalTime.Seconds;
                waveout.Init(afr);
                waveout.Play();
                PlayButton.Text = "播放";
                Total_Label.Text = afr.TotalTime.ToString().Substring(0, afr.TotalTime.ToString().LastIndexOf("."));
                if (!TimeLabelsBox.Items.Contains("00:00:00"))
                {
                    TimeLabelsBox.Items.Add("00:00:00");
                }
                if (!TimeLabelsBox.Items.Contains(Total_Label.Text))
                {
                    TimeLabelsBox.Items.Add(Total_Label.Text);
                }
                if (TimeLabelsBox.SelectedItems.Count == 0)
                {
                    TimeLabelsBox.SetSelected(0, true);
                }
                timer1.Start();
            }
        }

        private void StopButton_Click(object sender, EventArgs e)
        {
            if (waveout != null)
            {
                trackBar1.Value = 0;
                timer1.Stop();
                waveout.Stop();
                PlayButton.Text = "已停止";
                waveout.PlaybackStopped += OnPlaybackStopped;
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

        private void AddTimeLabel_Click(object sender, EventArgs e)
        {
            if (afr != null && waveout != null)
            {
                if (TimeLabelsBox.SelectedItems.Count == 0)
                {
                    TimeLabelsBox.Items.Add(Current_Label.Text);
                    TimeLabelsBox.SetSelected(TimeLabelsBox.Items.Count - 1, true);
                }
                else if (TimeLabelsBox.SelectedItems.Count == 1)
                {
                    int n = TimeLabelsBox.SelectedIndex;
                    TimeLabelsBox.Items.Insert(n + 1, Current_Label.Text);
                    TimeLabelsBox.SelectedItems.Clear();
                    TimeLabelsBox.SetSelected(n + 1, true);
                }
                else
                {
                    MessageBox.Show("请注意：标签将插入到第一个所选标签的后面，建议只选中一个标签然后进行新标签的插入");
                    int n = TimeLabelsBox.SelectedIndex;
                    TimeLabelsBox.Items.Insert(n + 1, Current_Label.Text);
                    TimeLabelsBox.SelectedItems.Clear();
                    TimeLabelsBox.SetSelected(n + 1, true);
                }
            }
            else
            {
                MessageBox.Show("请播放音频");
            }
        }

        private void DeleteLabelsMenu_Click(object sender, EventArgs e)
        {
            if (TimeLabelsBox.SelectedItems.Count != 0)
            {
                for (int i = TimeLabelsBox.SelectedItems.Count - 1; i >= 0; i--)
			    {
                    if (TimeLabelsBox.SelectedItems[i].ToString() == "00:00:00" || TimeLabelsBox.SelectedItems[i].ToString() == Total_Label.Text)
                    {
                        MessageBox.Show("不能删除首尾标签");
                    }
                    else
                    {
                        TimeLabelsBox.Items.Remove(TimeLabelsBox.SelectedItems[i]);
                    }
			    }
            }
        }

        private void DeleteAllLabelsMenu_Click(object sender, EventArgs e)
        {
            if (TimeLabelsBox.Items.Count != 0)
            {
                for (int i = TimeLabelsBox.Items.Count - 1; i >= 0; i--)
                {
                    if (TimeLabelsBox.Items[i].ToString() != "00:00:00" && TimeLabelsBox.Items[i].ToString() != Total_Label.Text)
                    {
                        TimeLabelsBox.Items.RemoveAt(i);
                    }      
                }
            }
        }

        private static int ChangeTimeFormat(string timelabel)
        {
            string[] arr = timelabel.Split(char.Parse(":")).ToArray();
            int h = int.Parse(arr[0]) * 3600;
            int m = int.Parse(arr[1]) * 60;
            int s = int.Parse(arr[2]);
            return h + m + s;
        }

        private static string ChangeTimeFormat2(int timelabel2)
        {
            int a = timelabel2;
            int h = a / 3600;
            int m = (a - h * 3600) / 60;
            int s = a - h * 3600 - m * 60;
            string nh = h.ToString();
            string nm = m.ToString();
            string ns = s.ToString();
            if (h < 10)
            {
                nh = "0" + h;
            }
            if (m < 10)
            {
                nm = "0" + m;
            }
            if (s < 10)
            {
                ns = "0" + s;
            }
            return nh + ":" + nm + ":" + ns;
        }

        private void SplitAudioButtonAll_Click(object sender, EventArgs e)
        {
            if (inputname != "")
            {
                if (TimeLabelsBox.Items.Count == 0)
                {
                    MessageBox.Show("请先添加时间标签");
                }
                else
                {
                    this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                    StopButton_Click(null, null);
                    if (waveout != null)
                    {
                        waveout.Dispose();
                        waveout = null;
                    }
                    if (afr != null)
                    {
                        afr.Dispose();
                        afr = null;
                    }
                    List<int> times = new List<int>();
                    for (int i = 0; i < TimeLabelsBox.Items.Count; i++)
                    {
                        times.Add(ChangeTimeFormat(TimeLabelsBox.Items[i].ToString()));
                    }
                    if (!times.Contains(0))
                    {
                        times.Add(0);
                    }
                    if (!times.Contains(ChangeTimeFormat(Total_Label.Text)))
                    {
                        times.Add(ChangeTimeFormat(Total_Label.Text));
                    }
                    times = times.Distinct().ToList();
                    times.Sort();

                    //Code Source http://stackoverflow.com/questions/6094287/naudio-to-split-mp3-file
                    string mp3dir = Path.GetDirectoryName(inputname);
                    string mp3name = Path.GetFileNameWithoutExtension(inputname);
                    string splitPath = Path.Combine(mp3dir, mp3name);
                    string extension = Path.GetExtension(inputname);
                    if (!Directory.Exists(splitPath))
                    {
                        Directory.CreateDirectory(splitPath);
                    }
                    DirectoryInfo dir = new System.IO.DirectoryInfo(splitPath);

                    if (extension == ".mp3")
                    {
                        for (int i = 0; i < times.Count() - 1; i++)
                        {
                            string outname = splitPath + "\\" + mp3name + "_" + dir.GetFiles().Length + ".mp3";
                            TrimMp3(inputname, outname, times[i], times[i + 1]);
                        }
                    }
                    else if (extension == ".wav")
                    {
                        for (int i = 0; i < times.Count() - 1; i++)
                        {
                            string outname = splitPath + "\\" + mp3name + "_" + dir.GetFiles().Length + ".wav";
                            TrimWav(inputname, outname, times[i], times[i + 1]);
                        }
                    }
                    this.Cursor = System.Windows.Forms.Cursors.Default;
                    System.Diagnostics.Process.Start("Explorer.exe", splitPath);
                } 
            }
            else
            {
                MessageBox.Show("请载入音频并添加时间标签");
            }
        }

        private void TrimMp3(string inputPath, string outputPath, int begin, int end)
        {
            using (var trimreader = new Mp3FileReader(inputPath))  //http://stackoverflow.com/questions/7932951/trimming-mp3-files-using-naudio/14169073#14169073
            {
                using (var trimwriter = File.Create(outputPath))
                {
                    trimreader.Skip(begin);
                    Mp3Frame frame;
                    while ((frame = trimreader.ReadNextFrame()) != null)
                    {
                        if ((int)trimreader.CurrentTime.TotalSeconds >= begin)
                        {
                            if ((int)trimreader.CurrentTime.TotalSeconds <= end)
                            {
                                trimwriter.Write(frame.RawData, 0, frame.RawData.Length);
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                }
            }
        }
        
        public static void TrimWav(string inputPath, string outputPath, int start, int end)
        {
            using (WaveFileReader wavreader = new WaveFileReader(inputPath))
            {
                using (WaveFileWriter writer = new WaveFileWriter(outputPath, wavreader.WaveFormat))
                {
                    int segement = wavreader.WaveFormat.AverageBytesPerSecond;
                    int startPosition =  start * segement;
                    startPosition = startPosition - startPosition % wavreader.WaveFormat.BlockAlign;
                    int endPosition = end * segement;
                    endPosition = endPosition - endPosition % wavreader.WaveFormat.BlockAlign;
                    TrimWav2(wavreader, writer, startPosition, endPosition);
                }
            }
        }

        private static void TrimWav2(WaveFileReader reader, WaveFileWriter writer, int startPosition, int endPosition)
        {
            reader.Position = startPosition;
            byte[] buffer = new byte[1024];
            while (reader.Position < endPosition)
            {
                int segment = (int)(endPosition - reader.Position);
                if (segment > 0)
                {
                    int bytesToRead = Math.Min(segment, buffer.Length);
                    int bytesRead = reader.Read(buffer, 0, bytesToRead);
                    if (bytesRead > 0)
                    {
                        writer.Write(buffer, 0, bytesRead);
                    }
                }
            }
        }

        private void CombinButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.Multiselect = true;
            of.Filter = "MP3音频(*.mp3)|*.mp3|WAV音频(*.wav)|*.wav";
            
            if (of.ShowDialog() == DialogResult.OK)
            {
                string[] fileNames = of.FileNames;
                if (fileNames.Count() > 1)
                {
                    string extension = Path.GetExtension(of.FileName);
                    SaveFileDialog sf = new SaveFileDialog();
                    if (extension == ".mp3")
                    {
                        sf.Filter = "MP3音频(*.mp3)|*.mp3";
                    }
                    else if (extension == ".wav")
                    {
                        sf.Filter = "WAV音频(*.wav)|*.wav";
                    }

                    sf.AddExtension = true;
                    sf.Title = "音频保存为";
                    if (sf.ShowDialog() == DialogResult.OK)
                    {
                        this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                        string fpath = System.IO.Path.GetDirectoryName(of.FileName);
                        string mp3File = Path.GetFileName(of.FileName);
                        DirectoryInfo dir = new System.IO.DirectoryInfo(fpath);
                        FileStream fstream = null;
                        fstream = File.Create(sf.FileName);

                        if (extension == ".mp3")
                        {
                            CombineMP3(fileNames, fstream);
                        }
                        else if (extension == ".wav")
                        {
                            CombineWAV(fileNames, fstream);
                        }

                        if (waveout != null)
                        {
                            waveout.Dispose();
                        }
                        fstream.Dispose();
                        inputname = sf.FileName;
                        TitleLabel.Text = Path.GetFileName(inputname);
                        this.PlayButton_Click(null, null);
                        PlayButton.Enabled = true;
                        StopButton.Enabled = true;
                        trackBar1.Enabled = true;
                        this.Cursor = System.Windows.Forms.Cursors.Default;
                    }
                }
                else
                {
                    MessageBox.Show("只有一个音频无需合并");
                }
            }
        }

        public static void CombineMP3(string[] InputMP3, Stream OutputMP3)
        {
            Mp3FileReader cbreader = null;
            foreach (string File in InputMP3)
            {
                cbreader = new Mp3FileReader(File);
                if ((OutputMP3.Position == 0) && (cbreader.Id3v2Tag != null))
                {
                    OutputMP3.Write(cbreader.Id3v2Tag.RawData, 0, cbreader.Id3v2Tag.RawData.Length);
                }
                Mp3Frame frame;
                while ((frame = cbreader.ReadNextFrame()) != null)
                {
                    OutputMP3.Write(frame.RawData, 0, frame.RawData.Length);
                }
                cbreader.Dispose();
            }
        }

        public static void CombineWAV(string[] InputWAV, Stream OutputWAV)
        {
            //http://stackoverflow.com/questions/6777340/how-to-join-2-or-more-wav-files-together-programatically
            byte[] buffer = new byte[1024];
            WaveFileWriter waveFileWriter = null;

            try
            {
                foreach (string wavFile in InputWAV)
                {
                    using (WaveFileReader reader = new WaveFileReader(wavFile))
                    {
                        if (waveFileWriter == null)
                        {
                            waveFileWriter = new WaveFileWriter(OutputWAV, reader.WaveFormat);
                        }
                        else
                        {
                            if (!reader.WaveFormat.Equals(waveFileWriter.WaveFormat))
                            {
                                throw new InvalidOperationException("Can't concatenate WAV Files that don't share the same format");
                            }
                        }

                        int read;
                        while ((read = reader.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            waveFileWriter.Write(buffer, 0, read);
                        }
                    }
                }
            }
            finally
            {
                if (waveFileWriter != null)
                {
                    waveFileWriter.Dispose();
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

        private void SplitAudioButtonSelected_Click(object sender, EventArgs e)
        {
            if (inputname != "")
            {
                if (TimeLabelsBox.Items.Count == 0)
                {
                    MessageBox.Show("请先添加时间标签");
                }
                else
                {
                    if (TimeLabelsBox.SelectedItems.Count > 1)
                    {
                        this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
                        StopButton_Click(null, null);
                        if (waveout != null)
                        {
                            waveout.Dispose();
                            waveout = null;
                        }
                        if (afr != null)
                        {
                            afr.Dispose();
                            afr = null;
                        }
                        List<int> times = new List<int>();
                        for (int i = 0; i < TimeLabelsBox.SelectedItems.Count; i++)
                        {
                            times.Add(ChangeTimeFormat(TimeLabelsBox.SelectedItems[i].ToString()));
                        }
                        times = times.Distinct().ToList();
                        times.Sort();

                        string mp3dir = Path.GetDirectoryName(inputname);
                        string mp3name = Path.GetFileNameWithoutExtension(inputname);
                        string splitPath = Path.Combine(mp3dir, mp3name);
                        string extension = Path.GetExtension(inputname);
                        if (!Directory.Exists(splitPath))
                        {
                            Directory.CreateDirectory(splitPath);
                        }
                        DirectoryInfo dir = new System.IO.DirectoryInfo(splitPath);

                        if (extension == ".mp3")
                        {
                            for (int i = 0; i < times.Count() - 1; i++)
                            {
                                string outname = splitPath + "\\" + mp3name + "_" + dir.GetFiles().Length + ".mp3";
                                TrimMp3(inputname, outname, times[i], times[i + 1]);
                            }
                        }
                        else if (extension == ".wav")
                        {
                            for (int i = 0; i < times.Count() - 1; i++)
                            {
                                string outname = splitPath + "\\" + mp3name + "_" + dir.GetFiles().Length + ".wav";
                                TrimWav(inputname, outname, times[i], times[i + 1]);
                            }
                        }
                        this.Cursor = System.Windows.Forms.Cursors.Default;
                        System.Diagnostics.Process.Start("Explorer.exe", splitPath);
                    }
                    else
                    {
                        MessageBox.Show("请选中至少两个时间标签，作为分割的起始和终止时间");
                    }
                }
            }
            else
            {
                MessageBox.Show("请载入音频并添加时间标签");
            }
        }

        private void 重新排序ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            List<int> Times = new List<int>();
            foreach (string item in TimeLabelsBox.Items)
            {
                Times.Add(ChangeTimeFormat(item));
            }
            Times = Times.Distinct().ToList();
            Times.Sort();
            TimeLabelsBox.Items.Clear();
            for (int i = 0; i < Times.Count(); i++)
            {
                TimeLabelsBox.Items.Add(ChangeTimeFormat2(Times[i]));
            }
            TimeLabelsBox.SetSelected(TimeLabelsBox.Items.Count - 2, true);
        }

        private void TimeLabelsBox_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (TimeLabelsBox.Items.Count != 0 && TimeLabelsBox.SelectedItems.Count != 0)
                {
                    if (afr != null)
                    {
                        waveout.Stop();
                        waveout.PlaybackStopped += OnPlaybackStopped;
                        afr = new AudioFileReader(inputname);
                        waveout = new WaveOut();
                        waveout.Init(afr);
                        waveout.Play();
                        afr.Skip(ChangeTimeFormat(TimeLabelsBox.SelectedItem.ToString()));
                    }
                }
            }
        }

    }
}
