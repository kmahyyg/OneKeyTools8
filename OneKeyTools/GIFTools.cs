using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Gif.Components;
using System.IO;
using System.Drawing.Imaging;

namespace OneKeyTools
{
    public partial class GIFTools : Form
    {
        public GIFTools()
        {
            InitializeComponent();
            this.Deactivate += new EventHandler(color1_Deactivate);
        }

        void color1_Deactivate(object sender, EventArgs e)
        {
            timer1.Enabled = false;
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

        private void button1_Click(object sender, EventArgs e)
        {
            GIFTools.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button60.Enabled = true;
        }

        PowerPoint.Application app = Globals.ThisAddIn.Application;
        // 感谢ok群群友 yuanyilvhua(QQ:4570848**)提供测试电脑，以修复之前存在的高分屏下的bug :-)
        private void button2_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides && sel.SlideRange.Count > 1)
            {
                PowerPoint.Slides slides = app.ActivePresentation.Slides;
                PowerPoint.SlideRange srange = sel.SlideRange;
                int scount = srange.Count;

                float wp = app.ActivePresentation.PageSetup.SlideWidth;
                float hp = app.ActivePresentation.PageSetup.SlideHeight;
                int pw = Properties.Settings.Default.pwidth;
                int h2 = (int)(pw * hp / wp);

                string name = app.ActivePresentation.Name;
                if (name.Contains(".pptx"))
                {
                    name = name.Replace(".pptx", "");
                }
                if (name.Contains(".ppt"))
                {
                    name = name.Replace(".ppt", "");
                }
                string cPath = app.ActivePresentation.Path + @"\" + name + @" 的GIF图\";
                if (!Directory.Exists(cPath))
                {
                    Directory.CreateDirectory(cPath);
                }
                System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(cPath);
                int k = dir.GetFiles().Length + 1;
                string gpath = cPath + name + "_" + k + ".gif";
                int time = int.Parse(textBox1.Text.Trim());

                int[] index = new int[srange.Count];
                for (int i = 1; i <= srange.Count; i++)
                {
                    index[i - 1] = srange[i].SlideIndex;
                }
                if (index[0] > index[1])
                {
                    Array.Sort(index);
                }
                sel.Unselect();
                slides.Range(index).Select();

                List<string> path = new List<string>();
                for (int i = 1; i <= scount; i++)
                {
                    PowerPoint.Slide nslide = slides[index[i - 1]];
                    string cPath2 = cPath + name + "_" + k + ".jpg";
                    nslide.Export(cPath2, "jpg");
                    path.Add(cPath2);
                    k = k + 1;
                }

                AnimatedGifEncoder gif = new AnimatedGifEncoder();
                gif.Start(gpath);
                gif.SetDelay(time);
                if (checkBox1.Checked)
                {
                    gif.SetRepeat(0);
                }
                else
                {
                    gif.SetRepeat(-1);
                }

                Bitmap bmp = null;
                Graphics g = null;

                for (int j = 0; j < scount; j++)
                {
                    Image spic = Image.FromFile(path[j]);
                    if (j == 0)
                    {
                        bmp = new Bitmap(spic.Width, spic.Height);
                        bmp.SetResolution(spic.HorizontalResolution, spic.VerticalResolution);
                        g = Graphics.FromImage(bmp);
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    }
                    g.Clear(panel1.BackColor);
                    g.DrawImage(spic, 0, 0);
                    gif.AddFrame(bmp);
                    spic.Dispose();
                    File.Delete(path[j]);
                }
                gif.Finish();
                g.Dispose();
                bmp.Dispose();

                if (!checkBox2.Checked)
                {
                    System.Diagnostics.Process.Start("Explorer.exe", cPath);
                }
                else
                {
                    int n = srange[1].SlideNumber;
                    PowerPoint.Shape nshape = slides[n].Shapes.AddPicture(gpath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, -pw, 0, pw, h2);
                    if (pw > app.ActivePresentation.PageSetup.SlideWidth)
                    {
                        nshape.LockAspectRatio = Office.MsoTriState.msoTrue;
                        nshape.Width = app.ActivePresentation.PageSetup.SlideWidth;
                        nshape.Left = 0;
                        nshape.Top = 0;
                    }
                    nshape.Select();
                    File.Delete(gpath);
                    if (dir.GetFiles().Length == 0)
                    {
                        Directory.Delete(cPath, true);
                    }
                }
            }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                int count = range.Count;
                if (count > 1)
                {
                    string name = app.ActivePresentation.Name;
                    if (name.Contains(".pptx"))
                    {
                        name = name.Replace(".pptx", "");
                    }
                    if (name.Contains(".ppt"))
                    {
                        name = name.Replace(".ppt", "");
                    }
                    string cPath = app.ActivePresentation.Path + @"\" + name + @" 的GIF图\";
                    if (!Directory.Exists(cPath))
                    {
                        Directory.CreateDirectory(cPath);
                    }
                    System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(cPath);
                    int k = dir.GetFiles().Length + 1;
                    int time = int.Parse(textBox1.Text.Trim());
                    string gpath = cPath + name + "_" + k + ".gif";

                    List<int> w = new List<int>();
                    List<int> h = new List<int>();
                    List<string> path = new List<string>();
                    Bitmap spic = null;
                    float xs = 0; float ys = 0;
                    for (int i = 1; i <= count; i++)
                    {
                        PowerPoint.Shape pic = range[i];
                        pic.Export(cPath + name + "_" + k + ".png", PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                        spic = new Bitmap(cPath + name + "_" + k + ".png");
                        w.Add(spic.Width);
                        h.Add(spic.Height);
                        if (xs == 0)
                        {
                            xs = spic.HorizontalResolution;
                            ys = spic.VerticalResolution;
                        }
                        spic.Dispose();
                        path.Add(cPath + name + "_" + k + ".png");
                        k = k + 1;
                    }
                    int wmax = w.Max();
                    int hmax = h.Max();
                    Bitmap bmp = new Bitmap(wmax, hmax);
                    bmp.SetResolution(xs, ys);
                    Graphics g = Graphics.FromImage(bmp);
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;

                    AnimatedGifEncoder gif = new AnimatedGifEncoder();
                    gif.Start(gpath);
                    gif.SetDelay(time);
                    if (checkBox1.Checked)
                    {
                        gif.SetRepeat(0);
                    }
                    else
                    {
                        gif.SetRepeat(-1);
                    }
                    if (checkBox4.Checked)
                    {
                        gif.SetQuality(300);
                        gif.SetTransparent(panel1.BackColor);    
                    }

                    for (int j = 0; j < count; j++)
                    {
                        spic = new Bitmap(path[j]);
                        g.Clear(panel1.BackColor);
                        g.DrawImage(spic, (bmp.Width - spic.Width) / 2, (bmp.Height - spic.Height) / 2);
                        gif.AddFrame(bmp);
                        spic.Dispose();
                        File.Delete(path[j]);
                    }
                    gif.Finish();
                    g.Dispose();
                    bmp.Dispose();

                    if (!checkBox2.Checked)
                    {
                        System.Diagnostics.Process.Start("Explorer.exe", cPath);
                    }
                    else
                    {
                        PowerPoint.Shape nshape = slide.Shapes.AddPicture(gpath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, -wmax, 0, wmax, hmax);
                        nshape.Select();
                        //File.Delete(cPath + name + "_" + k + ".gif");
                        //if (dir.GetFiles().Length == 0)
                        //{
                        //    Directory.Delete(cPath, true);
                        //}
                    }
                }
                else
                {
                    MessageBox.Show("请选中多张图片");
                }
            }
            else
            {
                MessageBox.Show("请选中多张图片或幻灯片页面");
            }
        }

        private void label5_Click(object sender, EventArgs e)
        {
            string name = app.ActivePresentation.Name;
            if (name.Contains(".pptx"))
            {
                name = name.Replace(".pptx", "");
            }
            if (name.Contains(".ppt"))
            {
                name = name.Replace(".ppt", "");
            }
            string cPath = app.ActivePresentation.Path + @"\" + name + @" 的GIF图\";
            if (!Directory.Exists(cPath))
            {
                MessageBox.Show("不存在拼图文件夹");
            }
            else
            {
                System.Diagnostics.Process.Start("Explorer.exe", cPath);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PowerPoint.Slide slide=app.ActiveWindow.View.Slide;
            string name = app.ActivePresentation.Name;
            if (name.Contains(".pptx"))
            {
                name = name.Replace(".pptx", "");
            }
            if (name.Contains(".ppt"))
            {
                name = name.Replace(".ppt", "");
            }
            string cPath = app.ActivePresentation.Path + @"\" + name + @" 的GIF图\";

            string gifFile = "";
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            if (!Directory.Exists(cPath))
            {
                openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }
            else
	        {
                openFileDialog1.InitialDirectory = cPath;
	        }
            openFileDialog1.Filter = "Image files (*.gif)|*.gif";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                gifFile = openFileDialog1.FileName;

                if (!Directory.Exists(cPath))
                {
                    Directory.CreateDirectory(cPath);
                }

                System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(cPath);
                int k = dir.GetFiles().Length + 1;

                GifDecoder gifDecoder = new GifDecoder();
                gifDecoder.Read(gifFile);
                int gcount = gifDecoder.GetFrameCount();

                float swidth = app.ActivePresentation.PageSetup.SlideWidth / gcount;
                float sheight = app.ActivePresentation.PageSetup.SlideHeight * 0.5f;

                Image pic = Image.FromFile(gifFile);
                int height = pic.Height;
                int width = pic.Width;
                pic.Dispose();
                for (int j = 0; j < gcount; j++)
                {
                    Image frame = gifDecoder.GetFrame(j);
                    frame.Save(cPath + name + "_" + k + "_" + (j + 1).ToString() + ".png", ImageFormat.Png);
                    if (checkBox3.Checked)
                    {
                        slide.Shapes.AddPicture(cPath + name + "_" + k + "_" + (j + 1).ToString() + ".png", Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, j * swidth, sheight, width, height);
                        File.Delete(cPath + name + "_" + k + "_" + (j + 1).ToString() + ".png");
                        if (dir.GetFiles().Length == 0)
                        {
                            Directory.Delete(cPath, true);
                        }
                    }
                }

                if (!checkBox3.Checked)
                {
                    System.Diagnostics.Process.Start("Explorer.exe", cPath);
                }
            }
        }

        private void label6_Click(object sender, EventArgs e)
        {
            string name = app.ActivePresentation.Name;
            if (name.Contains(".pptx"))
            {
                name = name.Replace(".pptx", "");
            }
            if (name.Contains(".ppt"))
            {
                name = name.Replace(".ppt", "");
            }
            string cPath = app.ActivePresentation.Path + @"\" + name + @" 的GIF图\";
            if (!Directory.Exists(cPath))
            {
                MessageBox.Show("不存在拼图文件夹");
            }
            else
            {
                Directory.Delete(cPath, true);
                MessageBox.Show("删除成功");
            }
        }

        private void panel1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                timer1.Enabled = true;
            }
            if (e.Button == MouseButtons.Right)
            {
                if (this.colorDialog1.ShowDialog() == DialogResult.OK)
                {
                    this.panel1.BackColor = colorDialog1.Color;
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Point p = new Point(MousePosition.X, MousePosition.Y);
            Bitmap shotImage = new Bitmap(1, 1);
            Graphics dc = Graphics.FromImage(shotImage);
            dc.CopyFromScreen(p, new Point(0, 0), shotImage.Size);
            Color c = shotImage.GetPixel(0, 0);
            panel1.BackColor = c;
        }
    }
}
