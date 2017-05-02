using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Drawing.Imaging;
using System.Threading;
using System.Diagnostics;

namespace OneKeyTools
{
    public partial class Whiteboard : Form
    {
        public Whiteboard()
        {
            InitializeComponent();
            this.Deactivate += new EventHandler(color1_Deactivate);
        }

        void color1_Deactivate(object sender, EventArgs e)
        {
            timer1.Enabled = false;
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
            toolStripMenuItem2.Text = this.Width + "," + this.Height;
        }

        private void 关闭ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
            if (this.BackgroundImage != null)
            {
                this.BackgroundImage.Dispose();
            }
        }

        string path = "";
        private void 载入ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string picFile = "";
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            openFileDialog1.Filter = "Image Files(*.JPG;*.PNG;*.BMP;*.GIF)|*.JPG;*.PNG;*.BMP;*.GIF";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                picFile = openFileDialog1.FileName;
                Bitmap bmp = new Bitmap(picFile);
                this.BackgroundImage = bmp;
                if (!设置宽高ToolStripMenuItem.Checked)
                {
                    this.BackgroundImageLayout = ImageLayout.Center;
                    this.Width = bmp.Width;
                    this.Height = bmp.Height;
                }
                path = picFile;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Point p = new Point(MousePosition.X, MousePosition.Y);
            Bitmap shotImage = new Bitmap(1, 1);
            Graphics dc = Graphics.FromImage(shotImage);
            dc.CopyFromScreen(p, new Point(0, 0), shotImage.Size);
            Color c = shotImage.GetPixel(0, 0);
            this.BackColor = c;
        }

        private void 设置宽高ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (设置宽高ToolStripMenuItem.Checked)
            {
                设置宽高ToolStripMenuItem.Checked = false;
                this.Width = this.BackgroundImage.Width;
                this.Height = this.BackgroundImage.Height;
            }
            else
            {
                设置宽高ToolStripMenuItem.Checked = true;
                string[] arr = toolStripMenuItem2.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
                int w = int.Parse(arr[0]);
                int h = int.Parse(arr[1]);
                this.Width = w;
                this.Height = h;
            }
        }

        private void 取色器ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.BackgroundImage = null;
            timer1.Enabled = true;
        }

        private void 颜色框ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.colorDialog1.ShowDialog() == DialogResult.OK)
            {
                this.BackgroundImage = null;
                this.BackColor = colorDialog1.Color;
            }
        }

        private void gif3_KeyDown(object sender, KeyEventArgs e)
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

        private void 透明度ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("【快捷键】Ctrl + ↑ 提升透明度/ ↓降低透明度");
        }

        private void gif3_Load(object sender, EventArgs e)
        {
            this.Width = Properties.Settings.Default.gwidth;
            this.Height = Properties.Settings.Default.gheight;
            this.Opacity = Properties.Settings.Default.opacity;
            string a = Properties.Settings.Default.picpath;
            if (a == "no")
            {
                this.BackgroundImage = null;
                this.BackColor = Properties.Settings.Default.color;
            }
            else
            {
                try
                {
                    this.BackgroundImage = new Bitmap(a);
                }
                catch
                {
                    this.BackgroundImage = null;
                    this.BackColor = Properties.Settings.Default.color;
                }
            }
        }

        private void toolStripMenuItem2_TextChanged(object sender, EventArgs e)
        {
            string[] arr = toolStripMenuItem2.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
            int w = int.Parse(arr[0]);
            int h = int.Parse(arr[1]);
            this.Width = w;
            this.Height = h;
        }

        private void 保存设置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.gwidth = this.Width;
            Properties.Settings.Default.gheight = this.Height;
            Properties.Settings.Default.opacity = this.Opacity;
            if (this.BackgroundImage == null)
            {
                Properties.Settings.Default.color = this.BackColor;
                Properties.Settings.Default.picpath = "no";
            }
            else
            {
                Properties.Settings.Default.picpath = path;
            }
            Properties.Settings.Default.Save();
        }

        private void 恢复默认ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.gwidth = 200;
            Properties.Settings.Default.gheight = 200;
            Properties.Settings.Default.opacity = 1;
            if (this.BackgroundImage == null)
            {
                Properties.Settings.Default.color = Color.FromArgb(255, 240, 240, 240) ;
            }
            else
            {
                Properties.Settings.Default.picpath = "no";
            }
            Properties.Settings.Default.Save();

            this.Width = 200;
            this.Height = 200;
            this.Opacity = 1;
            if (this.BackgroundImage != null)
            {
                this.BackgroundImage = null;
            }
            
            this.BackColor = Color.FromArgb(255, 240, 240, 240);
        }

        private void 复制ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.gwidth = this.Width;
            Properties.Settings.Default.gheight = this.Height;
            Properties.Settings.Default.opacity = this.Opacity;
            if (this.BackgroundImage == null)
            {
                Properties.Settings.Default.color = this.BackColor;
            }
            else
            {
                Properties.Settings.Default.picpath = path;
            }

            Whiteboard gif3 = null;
            if (gif3 == null || gif3.IsDisposed)
            {
                gif3 = new Whiteboard();
                IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
                NativeWindow win = NativeWindow.FromHandle(handle);
                gif3.Show();
            }
        }

        private void 新建ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Whiteboard gif3 = null;
            if (gif3 == null || gif3.IsDisposed)
            {
                gif3 = new Whiteboard();
                IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
                NativeWindow win = NativeWindow.FromHandle(handle);
                gif3.Show();
            }

            gif3.Width = 200;
            gif3.Height = 200;
            gif3.Opacity = 1;
            if (gif3.BackgroundImage != null)
            {
                gif3.BackgroundImage = null;
            }

            gif3.BackColor = Color.FromArgb(255, 240, 240, 240);
        }

    }
}
