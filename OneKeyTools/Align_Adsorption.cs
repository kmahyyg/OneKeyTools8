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
using Forms = System.Windows.Forms;

namespace OneKeyTools
{
    public partial class Align_Adsorption : Form
    {
        public Align_Adsorption()
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

        private float PtoCM(float p)
        {
            float cm = (float)(p * 2.54 / 72);
            return cm;
        }

        private float CMtoP(float cm)
        {
            float p = (float)(cm * 72 / 2.54);
            return p;
        }

        private void label1_Click(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            float swidth = app.ActivePresentation.PageSetup.SlideWidth;
            float sheight = app.ActivePresentation.PageSetup.SlideHeight;
            int cn = 0;
            int cn2 = 0;
            if (checkBox1.Checked)
            {
                foreach (PowerPoint.Shape item in slide.Shapes)
                {
                    if (item.Name == "lines1")
                    {
                        cn += 1;
                    }
                }
                if (cn == 0)
                {
                    PowerPoint.Shape line1 = slide.Shapes.AddLine(swidth * 0.25f, 0, swidth * 0.25f, sheight);
                    line1.Visible = Office.MsoTriState.msoTrue;
                    line1.Name = "lines1";
                    line1.Line.ForeColor.RGB = 0;
                    line1.Line.DashStyle = Office.MsoLineDashStyle.msoLineSolid;
                }
                else
                {
                    MessageBox.Show("已经添加了水平线");
                }    
            }
            if (checkBox2.Checked)
            {
                foreach (PowerPoint.Shape item in slide.Shapes)
                {
                    if (item.Name == "lines2")
                    {
                        cn2 += 1;
                    }
                }
                if (cn2 == 0)
                {
                    PowerPoint.Shape line2 = slide.Shapes.AddLine(0, sheight * 0.25f, swidth, sheight * 0.25f);
                    line2.Visible = Office.MsoTriState.msoTrue;
                    line2.Name = "lines2";
                    line2.Line.ForeColor.RGB = 0;
                    line2.Line.DashStyle = Office.MsoLineDashStyle.msoLineSolid;
                }
                else
                {
                    MessageBox.Show("已经添加了垂直线");
                }       
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            int count = slide.Shapes.Count;
            for (int i = count; i >= 1; i--)
            {
                PowerPoint.Shape line = slide.Shapes[i];
                if (line.Name == "lines1" || line.Name == "lines2")
                {
                    line.Delete();
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("请先选中一个形状");
            }
            else
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.Slides slides = app.ActivePresentation.Slides;
                PowerPoint.Shape shape0 = sel.ShapeRange[1];
                float l0 = shape0.Left;
                float t0 = shape0.Top;
                if (checkBox1.Checked && !checkBox2.Checked)
                {
                    float left = 0;
                    int cn = 0;
                    foreach (PowerPoint.Shape item in slide.Shapes)
                    {
                        if (item.Name == "lines1")
                        {
                            left = item.Left;
                            cn += 1;
                        }
                    }
                    if (cn != 0)
                    {
                        if (!checkBox3.Checked)
                        {
                            foreach (PowerPoint.Shape item in slide.Shapes)
                            {
                                if ((item.Name != "lines1"  && item.Name != "lines2"))
	                            {
                                    if (item.Left >= left && l0 >= left && item.Left - left <= l0 - left)
                                    {
                                        item.Left = left;
                                    }
                                    else if (item.Left >= left && l0 < left && item.Left - left <= left - (l0 + shape0.Width))
                                    {
                                        item.Left = left;
                                    }
                                    else if (item.Left < left && item.Left + item.Width / 2 >= left)
                                    {
                                        item.Left = left;
                                    }
                                    else if (item.Left < left && (item.Left + item.Width / 2 < left && item.Left + item.Width > left))
                                    {
                                        item.Left = left - item.Width;
                                    }
                                    else if ((item.Left < left && item.Left + item.Width < left) && (l0 > left && left - item.Left - item.Width <= l0 - left))
                                    {
                                        item.Left = left - item.Width;
                                    }
                                    else if ((item.Left < left && item.Left + item.Width < left) && (l0 < left && l0 + shape0.Width < left && left - item.Left - item.Width <= left - l0 - shape0.Width))
                                    {
                                        item.Left = left - item.Width;
                                    }
	                            }
                            }
                        }
                        else
                        {
                            foreach (PowerPoint.Slide nslide in slides)
                            {
                                foreach (PowerPoint.Shape item in nslide.Shapes)
                                {
                                    if (item.Name != "lines1"  && item.Name != "lines2")
                                    {
                                        if (item.Left >= left && l0 >= left && item.Left - left <= l0 - left)
                                        {
                                            item.Left = left;
                                        }
                                        else if (item.Left >= left && l0 < left && item.Left - left <= left - (l0 + shape0.Width))
                                        {
                                            item.Left = left;
                                        }
                                        else if (item.Left < left && item.Left + item.Width / 2 >= left)
                                        {
                                            item.Left = left;
                                        }
                                        else if (item.Left < left && (item.Left + item.Width / 2 < left && item.Left + item.Width > left))
                                        {
                                            item.Left = left - item.Width;
                                        }
                                        else if ((item.Left < left && item.Left + item.Width < left) && (l0 > left && left - item.Left - item.Width <= l0 - left))
                                        {
                                            item.Left = left - item.Width;
                                        }
                                        else if ((item.Left < left && item.Left + item.Width < left) && (l0 < left && l0 + shape0.Width < left && left - item.Left - item.Width <= left - l0 - shape0.Width))
                                        {
                                            item.Left = left - item.Width;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("请先添加水平线");
                    }
                }
                else if (checkBox2.Checked && !checkBox1.Checked)
                {
                    float top = 0;
                    float cn = 0;
                    foreach (PowerPoint.Shape item in slide.Shapes)
                    {
                        if (item.Name == "lines2")
                        {
                            top = item.Top;
                            cn += 1;
                        }
                    }
                    if (cn != 0)
                    {
                        if (!checkBox3.Checked)
                        {
                            foreach (PowerPoint.Shape item in slide.Shapes)
                            {
                                if (item.Name != "lines1"  && item.Name != "lines2")
                                {
                                    if (item.Top >= top && t0 >= top && item.Top - top <= t0 - top)
                                    {
                                        item.Top = top;
                                    }
                                    else if (item.Top >= top && t0 < top && item.Top - top <= top - (t0 + shape0.Height))
                                    {
                                        item.Top = top;
                                    }
                                    else if (item.Top < top && item.Top + item.Height / 2 >= top)
                                    {
                                        item.Top = top;
                                    }
                                    else if (item.Top < top && (item.Top + item.Height / 2 < top && item.Top + item.Height > top))
                                    {
                                        item.Top = top - item.Height;
                                    }
                                    else if ((item.Top < top && item.Top + item.Height < top) && (t0 > top && top - item.Top - item.Height <= t0 - top))
                                    {
                                        item.Top = top - item.Height;
                                    }
                                    else if ((item.Top < top && item.Top + item.Height < top) && (t0 < top && t0 + shape0.Height < top && top - item.Top - item.Height <= top - t0 - shape0.Height))
                                    {
                                        item.Top = top - item.Height;
                                    }
                                }
                            }
                        }
                        else
                        {
                            foreach (PowerPoint.Slide nslide in slides)
                            {
                                foreach (PowerPoint.Shape item in nslide.Shapes)
                                {
                                    if (item.Name != "lines1" && item.Name != "lines2")
                                    {
                                        if (item.Top >= top && t0 >= top && item.Top - top <= t0 - top)
                                        {
                                            item.Top = top;
                                        }
                                        else if (item.Top >= top && t0 < top && item.Top - top <= top - (t0 + shape0.Height))
                                        {
                                            item.Top = top;
                                        }
                                        else if (item.Top < top && item.Top + item.Height / 2 >= top)
                                        {
                                            item.Top = top;
                                        }
                                        else if (item.Top < top && (item.Top + item.Height / 2 < top && item.Top + item.Height > top))
                                        {
                                            item.Top = top - item.Height;
                                        }
                                        else if ((item.Top < top && item.Top + item.Height < top) && (t0 > top && top - item.Top - item.Height <= t0 - top))
                                        {
                                            item.Top = top - item.Height;
                                        }
                                        else if ((item.Top < top && item.Top + item.Height < top) && (t0 < top && t0 + shape0.Height < top && top - item.Top - item.Height <= top - t0 - shape0.Height))
                                        {
                                            item.Top = top - item.Height;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("请先添加垂直线");
                    }
                }
                else if (checkBox1.Checked && checkBox2.Checked)
                {
                    float left = 0;
                    float top = 0;
                    int cn = 0;
                    int cn2 = 0;
                    foreach (PowerPoint.Shape item in slide.Shapes)
                    {
                        if (item.Name == "lines1")
                        {
                            left = item.Left;
                            cn += 1;
                        }
                        else if (item.Name == "lines2")
                        {
                            top = item.Top;
                            cn2 += 1;
                        }
                    }
                    if (cn != 0 && cn2 != 0)
                    {
                        if (!checkBox3.Checked)
                        {
                            foreach (PowerPoint.Shape item in slide.Shapes)
                            {
                                if (item.Name != "lines1" && item.Name != "lines2")
                                {
                                    if ((l0 >= left && ((item.Left >= left && item.Left - left <= l0 - left) || (item.Left <= left && item.Left +item.Width /2 >= left))) && (t0 >= top && ((item.Top >= top && item.Top + item.Height >= top && item.Top - top <= t0 - top) || (item.Top <= top && item.Top + item.Height / 2 >= top))))
                                    {
                                        item.Left = left;
                                        item.Top = top;
                                    }
                                    else if ((l0 >= left && ((item.Left >= left && item.Left - left <= l0 - left) || (item.Left <= left && item.Left + item.Width / 2 >= left))) && (t0 <= top && item.Top <= top && (item.Top + item.Height < top && top - item.Top - item.Height <= top - t0 - shape0.Height || (item.Top + item.Height / 2 <= top && item.Top + item.Height >= top))))
                                    {
                                        item.Left = left;
                                        item.Top = top - item.Height;
                                    }
                                    else if ((l0 <= left && item.Left <= left && (left - item.Left <= left - l0 || (item.Left + item.Width / 2 <= left && item.Left + item.Width >= left))) && (t0 >= top && ((item.Top >= top && item.Top - top <= t0 - top) || (item.Top <= top && item.Top + item.Height / 2 >= top))))
                                    {
                                        item.Left = left - item.Width;
                                        item.Top = top;
                                    }
                                    else if ((l0 <= left && ((item.Left <= left && left - item.Left <= left - l0) || (item.Left > left && item.Left + item.Width / 2 <= left && item.Left + item.Width >= left))) && (t0 <= top && ((item.Top <= top && top - item.Top <= top - t0) || (item.Top >= top && item.Top + item.Height / 2 <= top && item.Top + item.Height >= top))))
                                    {
                                        item.Left = left - item.Width;
                                        item.Top = top - item.Height;
                                    }
                                }
                            }
                        }
                        else
                        {
                            foreach (PowerPoint.Slide nslide in slides)
                            {
                                foreach (PowerPoint.Shape item in nslide.Shapes)
                                {
                                    if (item.Name != "lines1" && item.Name != "lines2")
                                    {
                                        if ((l0 >= left && ((item.Left >= left && item.Left - left <= l0 - left) || (item.Left <= left && item.Left + item.Width / 2 >= left))) && (t0 >= top && ((item.Top >= top && item.Top + item.Height >= top && item.Top - top <= t0 - top) || (item.Top <= top && item.Top + item.Height / 2 >= top))))
                                        {
                                            item.Left = left;
                                            item.Top = top;
                                        }
                                        else if ((l0 >= left && ((item.Left >= left && item.Left - left <= l0 - left) || (item.Left <= left && item.Left + item.Width / 2 >= left))) && (t0 <= top && item.Top <= top && (item.Top + item.Height < top && top - item.Top - item.Height <= top - t0 - shape0.Height || (item.Top + item.Height / 2 <= top && item.Top + item.Height >= top))))
                                        {
                                            item.Left = left;
                                            item.Top = top - item.Height;
                                        }
                                        else if ((l0 <= left && item.Left <= left && (left - item.Left <= left - l0 || (item.Left + item.Width / 2 <= left && item.Left + item.Width >= left))) && (t0 >= top && ((item.Top >= top && item.Top - top <= t0 - top) || (item.Top <= top && item.Top + item.Height / 2 >= top))))
                                        {
                                            item.Left = left - item.Width;
                                            item.Top = top;
                                        }
                                        else if ((l0 <= left && ((item.Left <= left && left - item.Left <= left - l0) || (item.Left > left && item.Left + item.Width / 2 <= left && item.Left + item.Width >= left))) && (t0 <= top && ((item.Top <= top && top - item.Top <= top - t0) || (item.Top >= top && item.Top + item.Height / 2 <= top && item.Top + item.Height >= top))))
                                        {
                                            item.Left = left - item.Width;
                                            item.Top = top - item.Height;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else if (cn == 0 && cn2 != 0)
                    {
                        MessageBox.Show("请先添加水平线");
                    }
                    else if (cn2 == 0 && cn != 0)
                    {
                        MessageBox.Show("请先添加垂直线");
                    }
                    else
                    {
                        MessageBox.Show("请先添加水平线和垂直线");
                    }
                }
                else
                {
                    MessageBox.Show("请先勾选上方的水平线或垂直线");
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Align_Adsorption.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button135.Enabled = true;
        }

    }
}
