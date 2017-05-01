using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using forms = System.Windows.Forms;

namespace OneKeyTools
{
    public partial class Picture_Jigsaw : Form
    {
        public Picture_Jigsaw()
        {
            InitializeComponent();
            this.Deactivate += new EventHandler(color1_Deactivate);
        }

        void color1_Deactivate(object sender, EventArgs e)
        {
            timer1.Enabled = false;
        }

        private PowerPoint.Application app = Globals.ThisAddIn.Application;

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

        float w; float h;

        private void pintu1_Load(object sender, EventArgs e)
        {
            w = app.ActivePresentation.PageSetup.SlideWidth;
            h = app.ActivePresentation.PageSetup.SlideHeight;
            label10.Text = w + "*" + h;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            PowerPoint.Slides slides = app.ActivePresentation.Slides;
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                forms.MessageBox.Show("请选中要导出的幻灯片");
            }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
            {
                int wresolution = int.Parse(textBox1.Text.Trim());
                int prow = int.Parse(textBox2.Text.Trim());
                int pspac = int.Parse(textBox3.Text.Trim());
                int pspac2 = int.Parse(textBox4.Text.Trim());
                String[] arr = textBox5.Text.Trim().Split(char.Parse(","), char.Parse(" "), char.Parse("，")).ToArray();

                if (wresolution <= 0 || prow <= 0 || pspac < 0 || pspac2 < 0 || arr.Count() <= 0)
                {
                    MessageBox.Show("水平宽度/列数必须是正整数，间隔大小必须是非负整数,大图设置中的页码不能为空");
                }
                else
                {
                    PowerPoint.SlideRange srange = sel.SlideRange;
                    int count = srange.Count;
                    int[] index = new int[count];                      //index数组用于将页面的选择顺序强制从前到后
                    for (int i = 1; i <= count; i++)
                    {
                        index[i - 1] = srange[i].SlideIndex;
                    }
                    Array.Reverse(index);
                    Array.Sort(index);
                    sel.Unselect();
                    slides.Range(index).Select();

                    int[,] arrnew = new int[index.Count(), 2];             //arrnew用于初始化，arrnew[i,0]是所有页面的序号，arrnew[i,1]默认是小图
                    for (int i = 0; i < index.Count(); i++)
                    {
                        arrnew[i, 0] = index[i];
                        arrnew[i, 1] = 0;
                    }

                    if (arr.Count() > 1)                                    //将arr数组中的“n”替换为选择的最后一页的页码
                    {
                        for (int i = 0; i < arr.Count(); i++)
                        {
                            if (arr[i] == "n")
                            {
                                arr[i] = count.ToString();
                            }
                        }
                    }
                    else if (arr[0] == "n")
                    {
                        arr[0] = count.ToString();
                    }

                    int arrprow1 = 0;
                    for (int i = 0; i < arrnew.Length / 2; i++)                  //根据数组arr标记arrnew[i,1]中哪些页面是大图
                    {
                        if (prow == 1)
                        {
                            arrnew[i, 1] = 1;
                            arrprow1 += 1;
                        }
                        else
                        {
                            for (int j = 0; j < arr.Count(); j++)
                            {
                                if (arrnew[i, 0] == int.Parse(arr[j]))
                                {
                                    arrnew[i, 1] = 1;
                                    arrprow1 += 1;
                                }
                            }
                        }
                    }

                    int wlarge = wresolution - pspac2 * 2;         //根据用户设置的分辨率，计算大图和小图的宽度和高度
                    int hlarge = (int)(wlarge * h / w);
                    int wsmall = (int)((wlarge - (prow - 1) * pspac) / prow);
                    int hsmall = (int)(wsmall * h / w);

                    int arrcount = arrnew.Length / 2;
                    int[,] narr = new int[arrcount, 2];  //narr[0,0]是大图小图标识、narr[0,1]是水平序号
                    int wcount = 0;                      //wcount是水平序号、hcount是垂直序号、hscount是垂直方向上小图的行数、hscan是小图是否重新起一行
                    int hcount = 0;
                    int hscount = 0;
                    int hscan = 0;
                    for (int i = 0; i < arrcount; i++)
                    {
                        if (arrnew[i, 1] == 1)
                        {
                            narr[i, 0] = 1;
                            hcount += 1;
                            wcount = 0;
                            hscan = 0;
                        }
                        else
                        {
                            narr[i, 0] = 0;
                            if (wcount == 0)
	                        {
                                if (hscan == 0)
                                {
                                    narr[i, 1] = wcount;
                                    wcount += 1;
                                    hscount += 1;
                                    hcount += 1;
                                }
                                else
                                {
                                    wcount += 1;
                                    narr[i, 1] = wcount;
                                    wcount += 1;
                                }
	                        }
                            else
                            {
                                if (wcount < prow)
                                {
                                    narr[i, 1] = wcount;
                                    wcount += 1;
                                }
                                else
                                {
                                    wcount = 0;
                                    narr[i, 1] = wcount;
                                    hscount += 1;
                                    hcount += 1;
                                    hscan = 1;
                                }
                            }
                        }
                    }

                    Bitmap bmp0 = new Bitmap(wresolution, hlarge * arrprow1 + hsmall * hscount + pspac * (hcount - 1) + pspac2 * 2);    //计算长图的尺寸、设置长图的分辨率
                    float dpi = Properties.Settings.Default.dpi;
                    bmp0.SetResolution(dpi, dpi);

                    string name = app.ActivePresentation.Name;                  //根据演示文稿的文件名创建长图文件夹
                    if (name.Contains(".pptx"))
                    {
                        name = name.Replace(".pptx", "");
                    }
                    else if (name.Contains(".ppt"))
                    {
                        name = name.Replace(".ppt", "");
                    }
                    string cPath = app.ActivePresentation.Path + @"\" + name + @" 的幻灯片\";

                    if (!Directory.Exists(cPath))
                    {
                        Directory.CreateDirectory(cPath);
                    }

                    for (int i = 1; i <= sel.SlideRange.Count; i++)                //导出所选的页面为图片
                    {
                        PowerPoint.Slide nslide = sel.SlideRange[i];
                        string shname = name + "_临时_" + i;
                        nslide.Export(cPath + shname + ".jpg", "jpg", wresolution, (int)(wresolution * h / w));
                    }

                    Graphics g = Graphics.FromImage(bmp0);
                    if (!checkBox3.Checked)                                         //设置长图的底色
                    {
                        SolidBrush br = new SolidBrush(panel1.BackColor);
                        g.FillRectangle(br, 0, 0, bmp0.Width, bmp0.Height);
                    }

                    int ny = pspac2;
                    int sc = 1;
                    for (int i = 1; i <= count; i++)                                //读取之前导出的临时图片，根据之前的narr数组计算该图片所在的位置和尺寸
                    {
                        string shname2 = name + "_临时_" + i;
                        Bitmap bmp1 = new Bitmap(cPath + shname2 + ".jpg");
                        int x = 0;
                        int y = 0;
                        int wd = 0;
                        int ht = 0;

                        if (narr[i - 1, 0] == 1)
                        {
                            wd = wlarge;
                            ht = hlarge;
                            x = pspac2;
                            y = y + ny;
                            ny = ny + ht + pspac;
                        }
                        else
                        {
                            wd = wsmall;
                            ht = hsmall;
                            x = pspac2 + narr[i - 1, 1] * (wd + pspac);
                            y = y + ny;
                            if (sc < prow)
                            {
                                if (i < count && narr[i, 0] == 1)
                                {
                                    sc = 1;
                                    ny = ny + ht + pspac;
                                }
                                else
                                {
                                    sc += 1;
                                }
                            }
                            else
                            {
                                sc = 1;
                                ny = ny + ht + pspac;
                            }
                        }
                        g.DrawImage(bmp1, x, y, wd, ht);                          //在长图中进行绘制
                        bmp1.Dispose();
                        File.Delete(cPath + shname2 + ".jpg");
                    }

                    System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(cPath);   //保存长图为png或jpg
                    int k = dir.GetFiles().Length + 1;
                    if (checkBox3.Checked)
                    {
                        bmp0.Save(cPath + name + "_" + k + ".png", ImageFormat.Png);
                    }
                    else
                    {
                        bmp0.Save(cPath + name + "_" + k + ".jpg", ImageFormat.Jpeg);
                    }
                    bmp0.Dispose();
                    System.Diagnostics.Process.Start("Explorer.exe", cPath);            //完成后，打开长图所在文件夹
                }
            }
            else
            {
                MessageBox.Show("请选中至少1张幻灯片");
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

        private void button2_Click(object sender, EventArgs e)
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
            string cPath = app.ActivePresentation.Path + @"\" + name + @" 的幻灯片\";
            if (!Directory.Exists(cPath))
            {
                MessageBox.Show("不存在拼图文件夹");
            }
            else
            {
                System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(cPath);
                FileInfo[] files = dir.GetFiles();
                foreach (FileInfo file in files)
                { 
                    file.Delete(); 
                }
                MessageBox.Show("已清空拼图文件夹");
            }
        }

        private void button3_Click(object sender, EventArgs e)
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
            string cPath = app.ActivePresentation.Path + @"\" + name + @" 的幻灯片\";
            if (!Directory.Exists(cPath))
            {
                MessageBox.Show("不存在拼图文件夹");
            }
            else
            {
                Directory.Delete(cPath, true);
                MessageBox.Show("已删除拼图文件夹");
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (int.Parse(textBox2.Text.Trim()) <= 0)
            {
                MessageBox.Show("列数不能为0或负整数");
                textBox2.Text = "2";
            }

        }

        private void label10_Click(object sender, EventArgs e)
        {
            string[] w = label10.Text.Split(char.Parse("*")).ToArray();
            textBox1.Text = w[0];
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Picture_Jigsaw.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button170.Enabled = true;
        }

     }
}
