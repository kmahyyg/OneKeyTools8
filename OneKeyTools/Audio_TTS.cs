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
using System.Speech.Synthesis;
using System.IO;
using NAudio.Wave;

namespace OneKeyTools
{
    public partial class Audio_TTS : Form
    {
        public Audio_TTS()
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
            speech.Dispose();
            Audio_TTS.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button285.Enabled = true;
        }

        private static SpeechSynthesizer speech;
        private PromptBuilder pb;
        string content = "";
        string name = "";
        string name0 = "";
        string FileType = "mp3";
        string cPath = "";
        List<string> contents = null;
        List<int> contentsid = null;
        List<string> outnames = null;
        List<string> outnames2 = null;

        private void Audio_TTS_Load(object sender, EventArgs e)
        {
            speech = new SpeechSynthesizer();
            int n = speech.GetInstalledVoices().Count;
            EngineBox.Items.Clear();
            for (int i = 0; i < n; i++)
            {
                EngineBox.Items.Add(speech.GetInstalledVoices()[i].VoiceInfo.Name);
            }
            EngineBox.Text = EngineBox.Items[0].ToString();
            OutTypeBox.Text = "mp3";

            string pname = app.ActivePresentation.Name;
            if (pname.Contains(".pptx"))
            {
                pname = pname.Replace(".pptx", "");
            }
            if (pname.Contains(".ppt"))
            {
                pname = pname.Replace(".ppt", "");
            }
            cPath = app.ActivePresentation.Path + @"\" + pname + @" 的语音\";
        }

        private static void WavToMp3(string WavFile, string Mp3File)
        {
            NAudio.MediaFoundation.MediaFoundationApi.Startup();
            var mediaType = MediaFoundationEncoder.SelectMediaType(NAudio.MediaFoundation.AudioSubtypes.MFAudioFormat_WMAudioV8, new WaveFormat(16000, 1), 16000);
            if (mediaType != null)
            {
                using (var reader = new WaveFileReader(WavFile))
                {
                    MediaFoundationEncoder.EncodeToMp3(reader, Mp3File, 48000);
                }
            }
            NAudio.MediaFoundation.MediaFoundationApi.Shutdown();
        }

        private void SpeakButton_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            content = "";
            if (ReadNotes.Checked)
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    List<int> num = new List<int>();
                    foreach (PowerPoint.Slide slide in sel.SlideRange)
                    {
                        num.Add(slide.SlideNumber);
                    }
                    num.Sort();
                    for (int i = 0; i < num.Count(); i++)
                    {
                        PowerPoint.Slide slide = app.ActivePresentation.Slides[num[i]];
                        if (slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text != "")
                        {
                            content += slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text + "  ";
                        }
                    }
                }
                else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                    if (slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text != "")
                    {
                        content = slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text + "  ";
                    }
                }
            }
            else
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    if (sel.HasChildShapeRange)
                    {
                        range = sel.ChildShapeRange;
                    }

                    for (int i = 1; i <= range.Count; i++)
                    {
                        if (range[i].HasTextFrame == Office.MsoTriState.msoTrue && range[i].TextFrame.HasText == Office.MsoTriState.msoTrue && range[i].TextFrame.TextRange.Text != "")
                        {
                            content += range[i].TextFrame.TextRange.Text + "  ";
                        }
                    }
                }
                else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    List<int> num = new List<int>();
                    foreach (PowerPoint.Slide slide in sel.SlideRange)
                    {
                        num.Add(slide.SlideNumber);
                    }
                    num.Sort();
                    for (int i = 0; i < num.Count(); i++)
                    {
                        PowerPoint.Slide slide = app.ActivePresentation.Slides[num[i]];
                        foreach (PowerPoint.Shape shape in slide.Shapes)
                        {
                            if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue && shape.TextFrame.TextRange.Text != "")
                            {
                                content += shape.TextFrame.TextRange.Text + "  ";
                            }
                        }
                    }
                }
                else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    content = sel.TextRange.Text;
                }
            }
            
            if (content == "")
            {
                MessageBox.Show("没有找到文字，请确保选中带有文字的形状、幻灯片页面、有备注的页面");
            }
            else
            {
                try
                {
                    speech.Dispose();
                    speech = new SpeechSynthesizer();
                    speech.SetOutputToDefaultAudioDevice();
                    speech.SelectVoice(EngineBox.Text);
                    speech.Rate = trackBar1.Value;
                    speech.Volume = trackBar2.Value;
                    pb = new PromptBuilder();
                    pb.AppendText(content);
                    pb.Culture = speech.Voice.Culture;
                    speech.SpeakAsync(pb);
                }
                catch
                {
                    MessageBox.Show("发生错误，所选的声音包无法使用，请切换其他语音引擎");
                }
            } 
        }

        private void PauseButton_Click(object sender, EventArgs e)
        {
            if (speech != null && speech.State == SynthesizerState.Paused)
            {
                speech.Resume();
            }
            else if (speech != null && speech.State == SynthesizerState.Speaking)
            {
                speech.Pause();
            }
            else
            {
                MessageBox.Show("请先点朗读按钮");
            }
        }

        private void StopButton_Click(object sender, EventArgs e)
        {
            if (speech != null && (speech.State == SynthesizerState.Speaking || speech.State == SynthesizerState.Paused))
            {
                speech.SpeakAsyncCancelAll();
            }
            else
            {
                MessageBox.Show("请先点朗读按钮");
            }
        }

        private void OutTypeBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            FileType = OutTypeBox.Text;
        }

        private void OutputAllButton_Click(object sender, EventArgs e)
        {
            if (speech != null && (speech.State == SynthesizerState.Speaking || speech.State == SynthesizerState.Paused))
            {
                speech.SpeakAsyncCancelAll();
            }
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            content = "";
            if (ReadNotes.Checked)
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    List<int> num = new List<int>();
                    foreach (PowerPoint.Slide slide in sel.SlideRange)
                    {
                        num.Add(slide.SlideNumber);
                    }
                    num.Sort();
                    for (int i = 0; i < num.Count(); i++)
                    {
                        PowerPoint.Slide slide = app.ActivePresentation.Slides[num[i]];
                        if (slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text != "")
                        {
                            content += slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text + "  ";
                        }
                    }
                }
                else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                    if (slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text != "")
                    {
                        content = slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text + "  ";
                    }
                }
            }
            else
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    if (sel.HasChildShapeRange)
                    {
                        range = sel.ChildShapeRange;
                    }

                    for (int i = 1; i <= range.Count; i++)
                    {
                        if (range[i].HasTextFrame == Office.MsoTriState.msoTrue && range[i].TextFrame.HasText == Office.MsoTriState.msoTrue && range[i].TextFrame.TextRange.Text != "")
                        {
                            content += range[i].TextFrame.TextRange.Text + "  ";
                        }
                    }
                }
                else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                {
                    List<int> num = new List<int>();
                    foreach (PowerPoint.Slide slide in sel.SlideRange)
                    {
                        num.Add(slide.SlideNumber);
                    }
                    num.Sort();
                    for (int i = 0; i < num.Count(); i++)
                    {
                        PowerPoint.Slide slide = app.ActivePresentation.Slides[num[i]];
                        foreach (PowerPoint.Shape shape in slide.Shapes)
                        {
                            if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue && shape.TextFrame.TextRange.Text != "")
                            {
                                content += shape.TextFrame.TextRange.Text + "  ";
                            }
                        }
                    }
                }
                else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    content = sel.TextRange.Text;
                }
            }
            
            if (content == "")
            {
                MessageBox.Show("未找到文字，请确保选中带有文字的形状、幻灯片页面、有备注的页面");
            }
            else
            {
                try
                {
                    speech.Dispose();
                    speech = new SpeechSynthesizer();
                    speech.SelectVoice(EngineBox.Text);

                    if (!Directory.Exists(cPath))
                    {
                        Directory.CreateDirectory(cPath);
                    }
                    System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(cPath);
                    name0 = cPath + "合并语音_" + dir.GetFiles().Length;
                    name = name0 + ".wav";

                    speech.SetOutputToWaveFile(name);
                    speech.Rate = trackBar1.Value;
                    speech.Volume = trackBar2.Value;
                    pb = new PromptBuilder();

                    pb.AppendText(content);
                    pb.Culture = speech.Voice.Culture;
                    speech.SpeakAsync(pb);
                    speech.SpeakCompleted += new EventHandler<SpeakCompletedEventArgs>(speech_SpeakCompleted);
                }
                catch
                {
                    MessageBox.Show("发生错误，所选的声音包无法使用，请切换其他语音引擎");
                }
            }
            this.Cursor = System.Windows.Forms.Cursors.Default;
        }

        void speech_SpeakCompleted(object sender, SpeakCompletedEventArgs e)
        {
            if (FileType == "mp3")
            {
                speech.Dispose();
                speech = new SpeechSynthesizer();
                WavToMp3(name, name0 + ".mp3");
                File.Delete(name);
                name = name0 + ".mp3";
            }

            if (!IsToPPT.Checked)
            {
                System.Diagnostics.Process.Start("Explorer.exe", name);
            }
            else
            {
                PowerPoint.Slide cslide = app.ActiveWindow.View.Slide;
                PowerPoint.Shape audio = cslide.Shapes.AddMediaObject2(name, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 0, 0, 50, 50);
                audio.AnimationSettings.PlaySettings.HideWhileNotPlaying = Office.MsoTriState.msoTrue;
                audio.AnimationSettings.Animate = Office.MsoTriState.msoTrue;
                audio.AnimationSettings.AnimationOrder = 1;
                audio.AnimationSettings.AdvanceMode = PowerPoint.PpAdvanceMode.ppAdvanceOnTime;
            }
        }

        private void OutputSingleButton_Click(object sender, EventArgs e)
        {
            if (speech != null && (speech.State == SynthesizerState.Speaking || speech.State == SynthesizerState.Paused))
            {
                speech.SpeakAsyncCancelAll();
            }
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                MessageBox.Show("请选中幻灯片页面");
            }
            else
            {
                content = "";
                contents = new List<string>();
                contentsid = new List<int>();
                if (ReadNotes.Checked)
                {
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                    {
                        List<int> num = new List<int>();
                        foreach (PowerPoint.Slide slide in sel.SlideRange)
                        {
                            num.Add(slide.SlideNumber);
                        }
                        num.Sort();
                        for (int i = 0; i < num.Count(); i++)
                        {
                            PowerPoint.Slide slide = app.ActivePresentation.Slides[num[i]];
                            if (slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text != "")
                            {
                                contents.Add(slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text);
                                contentsid.Add(slide.SlideNumber);
                            }
                        }
                    }
                    else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                    {
                        PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                        if (slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text != "")
                        {
                            contents.Add(slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text);
                            contentsid.Add(slide.SlideNumber);
                        }
                    }
                }
                else
                {
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        if (sel.HasChildShapeRange)
                        {
                            range = sel.ChildShapeRange;
                        }

                        for (int i = 1; i <= range.Count; i++)
                        {
                            if (range[i].HasTextFrame == Office.MsoTriState.msoTrue && range[i].TextFrame.HasText == Office.MsoTriState.msoTrue && range[i].TextFrame.TextRange.Text != "")
                            {
                                contents.Add(range[i].TextFrame.TextRange.Text);
                                contentsid.Add(slide.SlideNumber);
                            }
                        }
                    }
                    else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                    {
                        PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                        contents.Add(sel.TextRange.Text);
                        contentsid.Add(slide.SlideNumber);
                    }
                    else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
                    {
                        List<int> num = new List<int>();
                        foreach (PowerPoint.Slide slide in sel.SlideRange)
                        {
                            num.Add(slide.SlideNumber);
                        }
                        num.Sort();
                        for (int i = 0; i < num.Count(); i++)
                        {
                            PowerPoint.Slide slide = app.ActivePresentation.Slides[num[i]];
                            foreach (PowerPoint.Shape shape in slide.Shapes)
                            {
                                if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue && shape.TextFrame.TextRange.Text != "")
                                {
                                    content += shape.TextFrame.TextRange.Text + "，";
                                }
                            }
                            if (content != "")
                            {
                                contents.Add(content);
                                contentsid.Add(slide.SlideNumber);
                                content = "";
                            }
                        }
                    }
                }
                
                if (contents.Count() != 0)
                {
                    if (!Directory.Exists(cPath))
                    {
                        Directory.CreateDirectory(cPath);
                    }
                    DirectoryInfo dir = new System.IO.DirectoryInfo(cPath);
                    outnames = new List<string>();
                    outnames2 = new List<string>();

                    try
                    {
                        for (int i = 0; i < contents.Count(); i++)
                        {
                            outnames.Add(cPath + "独立语音_" + dir.GetFiles().Length + ".wav");
                            outnames2.Add(cPath + "独立语音_" + dir.GetFiles().Length + "." + FileType);
                            speech = new SpeechSynthesizer();
                            speech.SelectVoice(EngineBox.Text);
                            speech.Rate = trackBar1.Value;
                            speech.Volume = trackBar2.Value;
                            speech.SetOutputToWaveFile(cPath + "独立语音_" + dir.GetFiles().Length + ".wav");
                            speech.Speak(contents[i]);
                            speech.Dispose();
                        }

                        speech.Dispose();
                        if (IsToPPT.Checked)
                        {
                            if (FileType == "mp3")
                            {
                                for (int i = 0; i < outnames.Count; i++)
                                {
                                    WavToMp3(outnames[i], outnames2[i]);
                                    File.Delete(outnames[i]);
                                    PowerPoint.Shape audio = app.ActivePresentation.Slides[contentsid[i]].Shapes.AddMediaObject2(outnames2[i], Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 0, 0, 50, 50);
                                    audio.AnimationSettings.PlaySettings.HideWhileNotPlaying = Office.MsoTriState.msoTrue;
                                    audio.AnimationSettings.Animate = Office.MsoTriState.msoTrue;
                                    audio.AnimationSettings.AnimationOrder = 1;
                                    audio.AnimationSettings.AdvanceMode = PowerPoint.PpAdvanceMode.ppAdvanceOnTime;
                                }
                            }
                            else
                            {
                                for (int i = 0; i < outnames.Count; i++)
                                {
                                    PowerPoint.Shape audio = app.ActivePresentation.Slides[contentsid[i]].Shapes.AddMediaObject2(outnames[i], Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 0, 0, 50, 50);
                                    audio.AnimationSettings.PlaySettings.HideWhileNotPlaying = Office.MsoTriState.msoTrue;
                                    audio.AnimationSettings.Animate = Office.MsoTriState.msoTrue;
                                    audio.AnimationSettings.AnimationOrder = 1;
                                    audio.AnimationSettings.AdvanceMode = PowerPoint.PpAdvanceMode.ppAdvanceOnTime;
                                }
                            }
                        }
                        else
                        {
                            if (FileType == "mp3")
                            {
                                for (int i = 0; i < outnames.Count; i++)
                                {
                                    WavToMp3(outnames[i], outnames2[i]);
                                    File.Delete(outnames[i]);
                                }
                            }
                            System.Diagnostics.Process.Start("Explorer.exe", cPath);
                        }
                        speech = new SpeechSynthesizer();
                    }
                    catch
                    {
                        MessageBox.Show("发生错误，所选的声音包无法使用，请切换其他语音引擎");
                    }        
                }
                else
                {
                    MessageBox.Show("所选页面没有文字");
                }
            }
            this.Cursor = System.Windows.Forms.Cursors.Default;
        }

        private void OutputFolder_Click(object sender, EventArgs e)
        {
            if (cPath != "")
            {
                System.Diagnostics.Process.Start("Explorer.exe", cPath);
            }
            else
            {
                MessageBox.Show("不存在语音文件夹，请关闭后重新打开本界面");
            }
        }
    }
}
