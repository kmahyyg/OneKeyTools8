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
    public partial class Picture_Wave : Form
    {
        public Picture_Wave()
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
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                string apath = app.ActivePresentation.Path;
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                int count = range.Count;
                int px = int.Parse(textBox1.Text.Trim());
                for (int i = 1; i <= count; i++)
                {
                    range[i].Export(apath + @"xshape-1.png", PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                    Bitmap bmp0 = new Bitmap(apath + @"xshape-1.png");
                    Bitmap bmp1 = null;
                    if (radioButton1.Checked)
                    {
                        bmp1 = new Bitmap(bmp0.Width, bmp0.Height + px);
                        int mw = bmp0.Width / 2;
                        float r = (float)(mw * mw + px * px) / (float)(2 * px);

                        for (int m = 0; m < bmp0.Width; m++)
                        {
                            for (int n = 0; n < bmp0.Height; n++)
                            {
                                Color color = bmp0.GetPixel(m, n);
                                int nn = 0;
                                nn = (int)(px + Math.Sqrt(r * r - (m - mw) * (m - mw)) - r);
                                bmp1.SetPixel(m, n + px - nn, color);
                            }
                        }
                    }
                    else if (radioButton2.Checked)
                    {
                        bmp1 = new Bitmap(bmp0.Width + px, bmp0.Height);
                        int mh = bmp0.Height / 2;
                        float r = (float)(mh * mh + px * px) / (float)(2 * px);

                        for (int m = 0; m < bmp0.Height; m++)
                        {
                            for (int n = 0; n < bmp0.Width; n++)
                            {
                                Color color = bmp0.GetPixel(n, m);
                                int nn = 0;
                                nn = (int)(px + Math.Sqrt(r * r - (m - mh) * (m - mh)) - r);
                                bmp1.SetPixel(n + nn, m, color);
                            }
                        }
                    }
                    bmp1.Save(apath + @"xshape-2.png");
                    slide.Shapes.AddPicture(apath + @"xshape-2.png", Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, app.ActivePresentation.PageSetup.SlideWidth/2,app.ActivePresentation.PageSetup.SlideHeight/2, bmp1.Width, bmp1.Height);
                    bmp1.Dispose();
                    bmp0.Dispose();
                    File.Delete(apath + @"xshape-1.png");
                    File.Delete(apath + @"xshape-2.png");
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Picture_Wave.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button205.Enabled = true;
        }
    }
}
