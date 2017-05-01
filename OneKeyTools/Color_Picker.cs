using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using forms = System.Windows.Forms;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace OneKeyTools
{
    public partial class Color_Picker : Form
    {
        public Color_Picker()
        {
            InitializeComponent();
            this.Deactivate += new EventHandler(color1_Deactivate);
        }

        void color1_Deactivate(object sender, EventArgs e)
        {
            timer1.Stop();
            timer2.Stop();
            timer3.Stop();
        }

        PowerPoint.Application app = Globals.ThisAddIn.Application;

        private int Rgb2Hsl(int r, int g, int b)
        {
            float h = 0;
            float s = 0;
            float l = 0;
            float max = Math.Max(Math.Max(r, g), b);
            float min = Math.Min(Math.Min(r, g), b);

            if (max == min)
            {
                h = 0;
            }
            else
            {
                if (max == r)
                {
                    if (g >= b)
                    {
                        h = 255 / 6 * (g - b) / (max - min) + 0;
                    }
                    else
                    {
                        h = 255 / 6 * (g - b) / (max - min) + 255;
                    }
                }
                if (max == g & max != r)
                {
                    h = 255 / 6 * (b - r) / (max - min) + 255 / 3;
                }
                if (max == b && max != g)
                {
                    h = 255 / 6 * (r - g) / (max - min) + 255 * 2 / 3;
                }
            }
            if (h >= (int)h + 0.5f)
            {
                h = (int)h + 1;
            }

            l = (max + min) / 2;
            if (max + min == 255)
            {
                l = 128;
            }
            if (l >= (int)l + 0.5f)
            {
                l = (int)l + 1;
            }

            if (l == 0 || max == min)
            {
                s = 0;
            }
            else
            {
                if (l <= 255 / 2)
                {
                    s = 255 * (max - min) / (max + min);
                }
                else
                {
                    s = 255 * (max - min) / (2 * 255 - (max + min));
                }
            }
            if (s >= (int)s + 0.5f)
            {
                s = (int)s + 1;
            }

            int hsl = (int)h + (int)s * 256 + (int)l * 256 * 256;
            return hsl;
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

        private void timer1_Tick(object sender, EventArgs e)
        {
            Point p = new Point(MousePosition.X, MousePosition.Y);
            Bitmap shotImage = new Bitmap(1, 1);
            Graphics dc = Graphics.FromImage(shotImage);
            dc.CopyFromScreen(p, new Point(0, 0), shotImage.Size);
            Color c = shotImage.GetPixel(0,0);
            int r = c.R;
            int g = c.G;
            int b = c.B;
            int hsl = Rgb2Hsl(r, g, b);
            int h = hsl % 256;
            int s = (hsl / 256) % 256;
            int l = (hsl / 256 / 256) % 256;
            label3.Text = string.Format("{0},{1},{2}", r, g, b);
            label4.Text = string.Format("{0},{1},{2}", h, s, l);
            label11.Text = ColorTranslator.ToHtml(c);

            panel1.BackColor = c;
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type!= PowerPoint.PpSelectionType.ppSelectionNone)
            {
                PowerPoint.ShapeRange range;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                else
                {
                    range = sel.ShapeRange;
                }
                int count = range.Count;
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape shape = range[i];
                    if (checkBox1.Checked)
                    {
                        shape.Fill.Visible = Office.MsoTriState.msoTrue;
                        shape.Fill.ForeColor.RGB = r + g * 256 + b * 256 * 256;                        
                    }
                    if (checkBox2.Checked)
                    {
                        shape.Line.Visible = Office.MsoTriState.msoTrue;
                        shape.Line.ForeColor.RGB = r + g * 256 + b * 256 * 256;
                    }
                    if (checkBox3.Checked)
                    {
                        shape.TextFrame.TextRange.Font.Color.RGB = r + g * 256 + b * 256 * 256;
                    }
                    if (checkBox5.Checked)
                    {
                        if (shape.Shadow.Visible == Office.MsoTriState.msoFalse)
                        {
                            shape.Shadow.Visible = Office.MsoTriState.msoTrue;
                        }
                        shape.Shadow.ForeColor.RGB = r + g * 256 + b * 256 * 256;
                    }
                }

            }
        }

        private void button1_Click(object sender, EventArgs e)
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

        private void panel2_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                timer2.Enabled = true;
            }
            if (e.Button == MouseButtons.Right)
            {
                if (this.colorDialog1.ShowDialog() == DialogResult.OK)
                {
                    this.panel2.BackColor = colorDialog1.Color;
                }
            }
        }

        private void panel3_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                timer3.Enabled = true;
            }
            if (e.Button == MouseButtons.Right)
            {
                if (this.colorDialog1.ShowDialog() == DialogResult.OK)
                {
                    this.panel3.BackColor = colorDialog1.Color;
                }
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            Point p = new Point(MousePosition.X, MousePosition.Y);
            Bitmap shotImage = new Bitmap(1, 1);
            Graphics dc = Graphics.FromImage(shotImage);
            dc.CopyFromScreen(p, new Point(0, 0), shotImage.Size); // 感谢PA插件开发者Pean的帮助
            Color c = shotImage.GetPixel(0, 0);
            int r = c.R; ;
            int g = c.G;
            int b = c.B;
            panel2.BackColor = Color.FromArgb(r, g, b);
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            Point p = new Point(MousePosition.X, MousePosition.Y);
            Bitmap shotImage = new Bitmap(1, 1);
            Graphics dc = Graphics.FromImage(shotImage);
            dc.CopyFromScreen(p, new Point(0, 0), shotImage.Size);
            Color c = shotImage.GetPixel(0, 0); 
            
            int r = c.R;
            int g = c.G;
            int b = c.B;
            panel3.BackColor = Color.FromArgb(r, g, b);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int r1 = panel2.BackColor.R;
            int g1 = panel2.BackColor.G;
            int b1 = panel2.BackColor.B;
            int r2 = panel3.BackColor.R;
            int g2 = panel3.BackColor.G;
            int b2 = panel3.BackColor.B;
            int rgb1 = r1 + g1 * 256 + b1 * 256 * 256;
            int rgb2 = r2 + g2 * 256 + b2 * 256 * 256;

            if (checkBox4.Checked)
            {
                PowerPoint.Slides slides = app.ActivePresentation.Slides;
                int scount = slides.Count;
                for (int i = 1; i <= scount; i++)
                {
                    int tcount = slides[i].Shapes.Count;
                    for (int j = 1; j <= tcount; j++)
                    {
                        PowerPoint.Shape shape = slides[i].Shapes[j];
                        if (shape.Type == Office.MsoShapeType.msoGroup)
                        {
                            foreach (PowerPoint.Shape item in shape.GroupItems)
                            {
                                if (checkBox1.Checked)
                                {
                                    if (item.Fill.Type == Office.MsoFillType.msoFillSolid && item.Fill.ForeColor.RGB == rgb1)
                                    {
                                        item.Fill.ForeColor.RGB = rgb2;
                                    }
                                }
                                if (checkBox2.Checked)
                                {
                                    if (item.Line.Visible == Office.MsoTriState.msoTrue && item.Line.Weight > 0 && item.Line.ForeColor.RGB == rgb1)
                                    {
                                        item.Line.ForeColor.RGB = rgb2;
                                    }
                                }
                                if (checkBox3.Checked)
                                {
                                    if (item.TextFrame.HasText == Office.MsoTriState.msoTrue && item.TextFrame.TextRange.Font.Color.RGB == rgb1)
                                    {
                                        item.TextFrame.TextRange.Font.Color.RGB = rgb2;
                                    }
                                    else if (item.TextFrame2.HasText == Office.MsoTriState.msoTrue && item.TextFrame2.TextRange.Font.Fill.ForeColor.RGB == rgb1)
                                    {
                                        item.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb2;
                                    }
                                }

                            }
                        }
                        else
	                    {
                            if (checkBox1.Checked)
                            {
                                if (shape.Fill.Type == Office.MsoFillType.msoFillSolid && shape.Fill.ForeColor.RGB == rgb1)
                                {
                                    shape.Fill.ForeColor.RGB = rgb2;
                                }
                            }
                            if (checkBox2.Checked)
                            {
                                if (shape.Line.Visible == Office.MsoTriState.msoTrue && shape.Line.Weight > 0 && shape.Line.ForeColor.RGB == rgb1)
                                {
                                    shape.Line.ForeColor.RGB = rgb2;
                                }
                            }
                            if (checkBox3.Checked)
                            {
                                if (shape.TextFrame.HasText == Office.MsoTriState.msoTrue && shape.TextFrame.TextRange.Font.Color.RGB == rgb1)
                                {
                                    shape.TextFrame.TextRange.Font.Color.RGB = rgb2;
                                }
                                else if (shape.TextFrame2.HasText == Office.MsoTriState.msoTrue && shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB == rgb1)
                                {
                                    shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb2;
                                }
                            }
	                    }                 
                    }
                }
            }
            else
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                int tcount = slide.Shapes.Count;
                for (int i = 1; i <= tcount; i++)
                {
                    PowerPoint.Shape shape = slide.Shapes[i];
                    if (shape.Type == Office.MsoShapeType.msoGroup)
                    {
                        foreach (PowerPoint.Shape item in shape.GroupItems)
                        {
                            if (checkBox1.Checked)
                            {
                                if (item.Fill.Type == Office.MsoFillType.msoFillSolid && item.Fill.ForeColor.RGB == rgb1)
                                {
                                    item.Fill.ForeColor.RGB = rgb2;
                                }
                            }
                            if (checkBox2.Checked)
                            {
                                if (item.Line.Visible == Office.MsoTriState.msoTrue && item.Line.Weight > 0 && item.Line.ForeColor.RGB == rgb1)
                                {
                                    item.Line.ForeColor.RGB = rgb2;
                                }
                            }
                            if (checkBox3.Checked)
                            {
                                if (item.TextFrame.HasText == Office.MsoTriState.msoTrue && item.TextFrame.TextRange.Font.Color.RGB == rgb1)
                                {
                                    item.TextFrame.TextRange.Font.Color.RGB = rgb2;
                                }
                                else if (item.TextFrame2.HasText == Office.MsoTriState.msoTrue && item.TextFrame2.TextRange.Font.Fill.ForeColor.RGB == rgb1)
                                {
                                    item.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb2;
                                }
                            }
                        }
                    }
                    else
                    {
                        if (checkBox1.Checked)
                        {
                            if (shape.Fill.Type == Office.MsoFillType.msoFillSolid && shape.Fill.ForeColor.RGB == rgb1)
                            {
                                shape.Fill.ForeColor.RGB = rgb2;
                            }
                        }
                        if (checkBox2.Checked)
                        {
                            if (shape.Line.Visible == Office.MsoTriState.msoTrue && shape.Line.Weight > 0 && shape.Line.ForeColor.RGB == rgb1)
                            {
                                shape.Line.ForeColor.RGB = rgb2;
                            }
                        }
                        if (checkBox3.Checked)
                        {
                            if (shape.TextFrame.HasText == Office.MsoTriState.msoTrue && shape.TextFrame.TextRange.Font.Color.RGB == rgb1)
                            {
                                shape.TextFrame.TextRange.Font.Color.RGB = rgb2;
                            }
                            else if (shape.TextFrame2.HasText == Office.MsoTriState.msoTrue && shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB == rgb1)
                            {
                                shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb2;
                            }
                        }
                    }
                }
            }
        }

        private void panel1_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionNone)
            {
                PowerPoint.ShapeRange range;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                else
                {
                    range = sel.ShapeRange;
                }
                int count = range.Count;
                int r = panel1.BackColor.R;
                int g = panel1.BackColor.G;
                int b = panel1.BackColor.B;
                int rgb = r + g * 256 + b * 256 * 256;
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape shape = range[i];
                    if (checkBox1.Checked)
                    {
                        shape.Fill.Visible = Office.MsoTriState.msoTrue;
                        shape.Fill.ForeColor.RGB = rgb;
                    }
                    if (checkBox2.Checked)
                    {
                        shape.Line.Visible = Office.MsoTriState.msoTrue;
                        shape.Line.ForeColor.RGB = rgb;
                    }
                    if (checkBox3.Checked)
                    {
                        shape.TextFrame.TextRange.Font.Color.RGB = rgb;
                    }
                    if (checkBox5.Checked)
                    {
                        if (shape.Shadow.Visible == Office.MsoTriState.msoFalse)
                        {
                            shape.Shadow.Visible = Office.MsoTriState.msoTrue;
                        }
                        shape.Shadow.ForeColor.RGB = r + g * 256 + b * 256 * 256;
                    }
                }
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(label3.Text);
        }

        private void label4_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(label4.Text);
        }

        private void label11_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(label11.Text);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Color_Picker.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button80.Enabled = true;
        }

    }
}
