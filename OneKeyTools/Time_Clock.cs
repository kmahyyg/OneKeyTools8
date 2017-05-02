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
    public partial class Time_Clock : Form
    {
        public Time_Clock()
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

        private void time1_Load(object sender, EventArgs e)
        {
            this.BackColor = Properties.Settings.Default.jsqbgcolor;
            label1.Font = Properties.Settings.Default.jsqfont;
            label1.ForeColor = Properties.Settings.Default.jsqfontcolor;
            Time_Clock.ActiveForm.Width = Properties.Settings.Default.twidth;
            Time_Clock.ActiveForm.Height = Properties.Settings.Default.theight;

            this.label1.Text = Convert.ToString(DateTime.Now.ToLocalTime());
            this.KeyPreview = true;
            label1.Left = (this.ClientRectangle.Width - label1.Width) / 2;
            label1.Top = (this.ClientRectangle.Height - label1.Height) / 2;
        }

        private void time1_Resize(object sender, System.EventArgs e)
        {
            label1.Left = (this.ClientRectangle.Width - label1.Width) / 2;
            label1.Top = (this.ClientRectangle.Height - label1.Height) / 2;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Label txt = this.label1;
            if (txt.Text.Length > 15 || txt.Text.Length == 0)
            {
                txt.Text = Convert.ToString(DateTime.Now.ToLocalTime());
            }
            else
            {
                txt.Text = DateTime.Now.ToLongTimeString();
            }
        }

        private void time1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                this.Opacity = this.Opacity - 0.10;
            }
            else if (e.KeyCode == Keys.Up)
            {
                this.Opacity = this.Opacity + 0.10;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                Time_Clock.ActiveForm.Close();
                Globals.Ribbons.Ribbon1.button25.Enabled = true;
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
                this.time1_Resize(sender, e);
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

        private void 关闭ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Time_Clock.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button25.Enabled = true;
        }

        private void 切换时钟格式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Label txt = this.label1;
            if (txt.Text.Length > 15)
            {
                txt.Text = DateTime.Now.ToLongTimeString();
            }
            else
            {
                txt.Text = Convert.ToString(DateTime.Now.ToLocalTime());
            }
            this.time1_Resize(sender, e);
        }

        private void 使用说明ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("1. 点击“切换显示格式”切换时间格式" + Environment.NewLine + "2. 点击“修改文字”和“修改背景”修改颜色或格式" + Environment.NewLine + "3. 单击窗体后，按上/下键可以降低/提高窗体透明度" + Environment.NewLine + "4. 点击“保存设置”，可保存除透明度与背景图片之外的所有设置" + Environment.NewLine + "5. 点击“恢复默认”恢复默认设置" + Environment.NewLine + "6. 单击窗体后，按【ESC】或右键点击关闭退出");
        }

        private void 保存设置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.twidth = this.ClientRectangle.Width;
            Properties.Settings.Default.theight = this.ClientRectangle.Height;
            Properties.Settings.Default.jsqfont = this.label1.Font;
            Properties.Settings.Default.jsqfontcolor = this.label1.ForeColor;
            Properties.Settings.Default.jsqbgcolor = this.BackColor;
            Properties.Settings.Default.Save();
            MessageBox.Show("已保存除透明度和背景图片之外的所有设置；" + Environment.NewLine + "本设置同样适用于“定时器”“倒计时”功能");
        }

        private void 恢复默认ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.jsqbgcolor = Color.FromArgb(255, 224, 67, 67);
            Properties.Settings.Default.jsqfontcolor = Color.White;
            Properties.Settings.Default.jsqfont = new Font("微软雅黑", 14.25f, FontStyle.Bold);
            Properties.Settings.Default.twidth = 300;
            Properties.Settings.Default.theight = 52;
            Properties.Settings.Default.jsqcan = 1;
            Properties.Settings.Default.Save();

            if (Properties.Settings.Default.jsqcan == 1)
            {
                Properties.Settings.Default.jsqcan = 0;
                Properties.Settings.Default.Save();
                MessageBox.Show("恢复成功，请重新打开数字时钟");
            }
            else
            {
                MessageBox.Show("恢复失败");
            }
        }
    }
}
