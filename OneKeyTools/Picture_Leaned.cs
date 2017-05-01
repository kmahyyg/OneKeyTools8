using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneKeyTools
{
    public partial class Picture_Leaned : Form
    {
        public Picture_Leaned()
        {
            InitializeComponent();
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

        private void button2_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("请先选中图片");
            }
            else
            {
                string apath = app.ActivePresentation.Path;
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                int count = range.Count;
                float nn = float.Parse(textBox1.Text.Trim()) / 100f;
                for (int p = 1; p <= count; p++)
                {
                    PowerPoint.Shape pic = range[p];
                    pic.Copy();
                    PowerPoint.Shape npic = slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
                    pic.Export(apath + @"xshape.png", PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                    Bitmap bmp0 = new Bitmap(apath + @"xshape.png");
                    Graphics g = Graphics.FromImage(bmp0);
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                    g.DrawImage(bmp0, 0, 0);
                    g.Dispose();

                    Bitmap bmp1 = null;
                    if (radioButton1.Checked)
                    {
                        bmp1 = new Bitmap(bmp0.Width, (int)(bmp0.Height * (Math.Abs(nn) + 1)));
                        for (int i = 0; i < bmp0.Width; i++)
                        {
                            for (int j = 0; j < bmp0.Height; j++)
                            {
                                Color color = bmp0.GetPixel(i, j);
                                double yn = 0;
                                if (nn >= 0)
                                {
                                    yn = (double)j + (double)Math.Abs(bmp0.Height * nn) * (double)i / (double)bmp0.Width;
                                }
                                else
                                {
                                    yn = (double)j + (double)Math.Abs(bmp0.Height * nn) * (double)(bmp0.Width - i) / (double)bmp0.Width;
                                }

                                bmp1.SetPixel(i, (int)yn, color);
                            }
                        }
                    }
                    else if (radioButton2.Checked)
                    {
                        bmp1 = new Bitmap((int)(bmp0.Width * (Math.Abs(nn) + 1)), bmp0.Height);

                        for (int i = 0; i < bmp0.Height; i++)
                        {
                            for (int j = 0; j < bmp0.Width; j++)
                            {
                                Color color = bmp0.GetPixel(j, i);
                                double xn = 0;
                                if (nn < 0)
                                {
                                    xn = (double)j + (double)Math.Abs(bmp0.Width * nn) * (double)i / (double)bmp0.Height;
                                }
                                else
                                {
                                    xn = (double)j + (double)Math.Abs(bmp0.Width * nn) * (double)(bmp0.Height - i) / (double)bmp0.Height;
                                }
                                bmp1.SetPixel((int)xn, i, color);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("先选择垂直或水平方向");
                    }
                    bmp1.Save(apath + @"xshape2.png");
                    PowerPoint.Shape pic2 = slide.Shapes.AddPicture(apath + @"xshape2.png", Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, pic.Left, pic.Top + pic.Height / 2 - npic.Height / 2, npic.Width, npic.Height);
                    npic.Delete();
                    bmp0.Dispose();
                    bmp1.Dispose();
                    File.Delete(apath + @"xshape.png");
                    File.Delete(apath + @"xshape2.png");
                    pic.Delete();
                    pic2.Select();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Picture_Leaned.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button216.Enabled = true;
        }
    }
}
