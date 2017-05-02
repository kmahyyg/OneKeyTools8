using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneKeyTools
{
    public partial class Time_CountDown : Form
    {
        public Time_CountDown()
        {
            InitializeComponent();
        }

        const int Guying_HTLEFT = 10; //无边框窗体拖动及调整大小代码，来自网络
        const int Guying_HTRIGHT = 11;
        const int Guying_HTTOP = 12;
        const int Guying_HTTOPLEFT = 13;
        const int Guying_HTTOPRIGHT = 14;
        const int Guying_HTBOTTOM = 15;
        const int Guying_HTBOTTOMLEFT = 0x10;
        const int Guying_HTBOTTOMRIGHT = 17;
        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case 0x0084:
                    base.WndProc(ref m);
                    Point vPoint = new Point((int)m.LParam & 0xFFFF,
                    (int)m.LParam >> 16 & 0xFFFF);
                    vPoint = PointToClient(vPoint);
                    if (vPoint.X <= 5)
                        if (vPoint.Y <= 5)
                            m.Result = (IntPtr)Guying_HTTOPLEFT;
                        else if (vPoint.Y >= ClientSize.Height - 5)
                            m.Result = (IntPtr)Guying_HTBOTTOMLEFT;
                        else m.Result = (IntPtr)Guying_HTLEFT;
                    else if (vPoint.X >= ClientSize.Width - 5)
                        if (vPoint.Y <= 5)
                            m.Result = (IntPtr)Guying_HTTOPRIGHT;
                        else if (vPoint.Y >= ClientSize.Height - 5)
                            m.Result = (IntPtr)Guying_HTBOTTOMRIGHT;
                        else m.Result = (IntPtr)Guying_HTRIGHT;
                    else if (vPoint.Y <= 5)
                        m.Result = (IntPtr)Guying_HTTOP;
                    else if (vPoint.Y >= ClientSize.Height - 5)
                        m.Result = (IntPtr)Guying_HTBOTTOM;
                    break;
                case 0x0201: //鼠标左键按下的消息
                    m.Msg = 0x00A1; //更改消息为非客户区按下鼠标
                    m.LParam = IntPtr.Zero; //默认值
                    m.WParam = new IntPtr(2);//鼠标放在标题栏内
                    base.WndProc(ref m);
                    break;
                default:
                    base.WndProc(ref m);
                    break;
            }
        }

        PowerPoint.Application app = Globals.ThisAddIn.Application;

        string jsqsz;
        private void Time_CountDown_Load(object sender, EventArgs e)
        {
            this.BackColor = Properties.Settings.Default.jsqbgcolor;
            label1.Font = Properties.Settings.Default.jsqfont;
            label1.ForeColor = Properties.Settings.Default.jsqfontcolor;
            this.Width = Properties.Settings.Default.twidth;
            this.Height = Properties.Settings.Default.theight;
            times = Properties.Settings.Default.jsqtime;
            jsqsz = Properties.Settings.Default.jsqsz;
            toolStripMenuItem2.Text = times;
            提示ToolStripMenuItem.Text = jsqsz;
            label1.Text = times;

            this.KeyPreview = true;
            label1.Left = (this.ClientRectangle.Width - label1.Width) / 2;
            label1.Top = (this.ClientRectangle.Height - label1.Height) / 2;
            app.SlideShowBegin += app_SlideShowBegin;
            app.SlideShowEnd += app_SlideShowEnd;
        }

        void app_SlideShowBegin(PowerPoint.SlideShowWindow wn)
        {
            timer1.Enabled = true;
        }

        void app_SlideShowEnd(PowerPoint.Presentation Pres)
        {
            timer1.Enabled = false;
        }

        private void Time_CountDown_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                this.Opacity = this.Opacity - 0.10;
            }
            else
            {
                if (e.KeyCode == Keys.Up)
                {
                    this.Opacity = this.Opacity + 0.10;
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (label1.Text.Contains(":"))
            {
                string[] arr = label1.Text.Split(char.Parse(":")).ToArray();
                int h = int.Parse(arr[0]) * 3600;
                int m = int.Parse(arr[1]) * 60;
                int s = int.Parse(arr[2]);
                int a = h + m + s - 1;
                if (a >= 0)
                {
                    h = a / 3600;
                    m = (a - h * 3600) / 60;
                    s = a - h * 3600 - m * 60;
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
                    this.label1.Text = nh + ":" + nm + ":" + ns;
                }
                else
                {
                    timer1.Enabled = false;
                    TipSound();
                    MessageBox.Show(jsqsz);
                }         
            }
            else
            {
                int a = int.Parse(label1.Text) - 1;
                if (a < 0)
                {
                    timer1.Enabled = false;
                    TipSound();
                    MessageBox.Show(jsqsz);
                }
                else
                {
                    label1.Text = a.ToString();
                }   
            }
        }

        private void label1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (timer1.Enabled == false)
                {
                    timer1.Enabled = true;
                }
                else
                {
                    timer1.Enabled = false;
                }
            }
        }

        private void 关闭ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Time_CountDown.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button154.Enabled = true;
        }

        private void Time_CountDown_Resize(object sender, EventArgs e)
        {
            label1.Left = (this.ClientRectangle.Width - label1.Width) / 2;
            label1.Top = (this.ClientRectangle.Height - label1.Height) / 2;
        }

        private void 重新开始ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            label1.Text = times;
        }

        private void 切换显示格式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (label1.Text.Contains(":"))
            {
                string[] arr = label1.Text.Split(char.Parse(":")).ToArray();
                int h = int.Parse(arr[0]) * 3600;
                int m = int.Parse(arr[1]) * 60;
                int s = int.Parse(arr[2]);
                this.label1.Text = Convert.ToString(h + m + s);
            }
            else
            {
                int a = int.Parse(label1.Text);
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
                this.label1.Text = nh + ":" + nm + ":" + ns;
            }
            this.Time_CountDown_Resize(sender, e);
        }

        string times;
        private void toolStripMenuItem2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == System.Convert.ToChar(13))
            {
                int time = int.Parse(toolStripMenuItem2.Text.Trim());
                if (time > 0)
                {
                    e.Handled = true;
                    timer1.Enabled = false;
                    times = time.ToString();
                    label1.Text = time.ToString();
                    MessageBox.Show("设置成功，左键单击时间开始倒计时");
                }
            }
        }

        private void 文字颜色ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.colorDialog1.ShowDialog() == DialogResult.OK)
            {
                this.label1.ForeColor = colorDialog1.Color;
            }
        }

        private void 文字格式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.fontDialog1.ShowDialog() == DialogResult.OK)
            {
                this.label1.Font = fontDialog1.Font;
                this.Time_CountDown_Resize(sender, e);
            }
        }

        private void 背景色ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.colorDialog1.ShowDialog() == DialogResult.OK)
            {
                this.BackColor = colorDialog1.Color;
                this.BackgroundImage = null;
            }
        }

        private void 背景图片ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.Filter = "jpg(*.jpg)|*.jpg|All files (*.*)|*.*";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string fName = openFileDialog.FileName;
                this.BackgroundImage = Image.FromFile(fName);
            }
        }

        private void 使用说明ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("1. 点“重新开始”重新开始倒计时" + Environment.NewLine + "2. 点击“切换显示格式”切换时间格式" + Environment.NewLine + "3. 在“设置时间”中输入倒计时开始时间（秒）并回车；在“设置提示”中修改文字和声音提示" + Environment.NewLine + "4. 点击“修改文字”和“修改背景”选择修改颜色或格式" + Environment.NewLine + "5. 单击窗体后，按上/下键可以降低/提高窗体透明度" + Environment.NewLine + "6. 点击“保存设置”，可保存除透明度与背景图片之外的所有设置" + Environment.NewLine + "7. 点击“恢复默认”恢复所有设置为默认" + Environment.NewLine + "8. 单击窗体后，按【ESC】或右键点击“关闭”退出");
        }

        private void 保存设置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.twidth = this.ClientRectangle.Width;
            Properties.Settings.Default.theight = this.ClientRectangle.Height;
            Properties.Settings.Default.jsqfont = this.label1.Font;
            Properties.Settings.Default.jsqfontcolor = this.label1.ForeColor;
            Properties.Settings.Default.jsqbgcolor = this.BackColor;
            Properties.Settings.Default.jsqsz = jsqsz;
            Properties.Settings.Default.jsqtime = times;
            Properties.Settings.Default.Save();
            MessageBox.Show("已保存除透明度和背景图片之外的所有设置；" + Environment.NewLine + "本设置同样适用于“数字时钟”“定时器”功能");
        }

        private void 提示ToolStripMenuItem_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == System.Convert.ToChar(13))
            {
                jsqsz = 提示ToolStripMenuItem.Text.Trim();
                MessageBox.Show("成功修改提示文字");
            }
        }

        private void 恢复默认ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.jsqbgcolor = Color.FromArgb(255, 224, 67, 67);
            Properties.Settings.Default.jsqfontcolor = Color.White;
            Properties.Settings.Default.jsqfont = new Font("微软雅黑", 14.25f, FontStyle.Bold);
            Properties.Settings.Default.twidth = 300;
            Properties.Settings.Default.theight = 52;
            Properties.Settings.Default.jsqtime = "60";
            Properties.Settings.Default.jsqsz = "时间到";
            Properties.Settings.Default.jsqcan = 1;
            Properties.Settings.Default.Save();

            if (Properties.Settings.Default.jsqcan == 1)
            {
                Properties.Settings.Default.jsqcan = 0;
                Properties.Settings.Default.Save();
                MessageBox.Show("恢复成功，请重新打开定时器");
            }
            else
            {
                MessageBox.Show("恢复失败");
            }
        }

        private void 设置声音ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Location = "";
            if (Properties.Settings.Default.TimeSound == "")
            {
                Microsoft.Win32.RegistryKey path = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Slibe\OneKeyTools", false);
                Location = path.GetValue("Path", "").ToString();
            }
            else
            {
                Location = Properties.Settings.Default.TimeSound;
            }
            OpenFileDialog of = new OpenFileDialog();
            of.InitialDirectory = Location;
            of.Filter = "WAV音频(*.wav)|*.wav|MP3音频(*.mp3)|*.mp3";

            if (of.ShowDialog() == DialogResult.OK)
            {
                Properties.Settings.Default.TimeSound = of.FileName;
                Properties.Settings.Default.Save();
                MessageBox.Show("声音提示设置成功");
            }
        }

        private AudioFileReader afr = null;
        private WaveOut waveout = null;

        private void TipSound()
        {
            if (Properties.Settings.Default.TimeSound != "")
            {
                if (afr != null)
                {
                    afr.Dispose();
                }
                if (waveout != null)
                {
                    waveout.Dispose();
                }
                afr = new AudioFileReader(Properties.Settings.Default.TimeSound);
                waveout = new WaveOut();
                waveout.Init(afr);
                waveout.Play();
            }
        }

    }
}
