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

namespace OneKeyTools
{
    public partial class Table_Color : Form
    {
        public Table_Color()
        {
            InitializeComponent();
            this.Deactivate += new EventHandler(color1_Deactivate);
        }

        PowerPoint.Application app = Globals.ThisAddIn.Application;

        void color1_Deactivate(object sender, EventArgs e)
        {
            timer1.Stop();
            timer2.Stop();
            timer3.Stop();
            timer4.Stop();
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

        private int Hsl2Rgb(int h, int s, int l)
        {
            float nh = (float)h / 255;
            float ns = (float)s / 255;
            float nl = (float)l / 255;
            float nr = 0;
            float ng = 0;
            float nb = 0;

            if (ns == 0)
            {
                nr = nl;
                ng = nl;
                nb = nl;
            }
            else
            {
                float q = 0;
                if (nl < 0.5f)
                {
                    q = nl * (1 + ns);
                }
                else
                {
                    q = nl + ns - nl * ns;
                }

                float p = 2 * nl - q;
                float tr = nh + (float)1 / 3;
                float tg = nh;
                float tb = nh - (float)1 / 3;

                if (tr < 0)
                {
                    tr = tr + 1;
                }
                else
                {
                    if (tr > 1)
                    {
                        tr = tr - 1;
                    }
                    else
                    {
                        tr = tr + 0;
                    }
                }

                if (tg < 0)
                {
                    tg = tg + 1;
                }
                else
                {
                    if (tg > 1)
                    {
                        tg = tg - 1;
                    }
                    else
                    {
                        tg = tg + 0;
                    }
                }

                if (tb < 0)
                {
                    tb = tb + 1;
                }
                else
                {
                    if (tb > 1)
                    {
                        tb = tb - 1;
                    }
                    else
                    {
                        tb = tb + 0;
                    }
                }

                if (tr < (float)1 / 6)
                {
                    nr = p + (q - p) * 6 * tr;
                }
                else
                {
                    if (tr < (float)1 / 2)
                    {
                        nr = q;
                    }
                    else
                    {
                        if (tr < (float)2 / 3)
                        {
                            nr = p + (q - p) * 6 * ((float)2 / 3 - tr);
                        }
                        else
                        {
                            nr = p;
                        }
                    }
                }

                if (tg < (float)1 / 6)
                {
                    ng = p + (q - p) * 6 * tg;
                }
                else
                {
                    if (tg < (float)1 / 2)
                    {
                        ng = q;
                    }
                    else
                    {
                        if (tg < (float)2 / 3)
                        {
                            ng = p + (q - p) * 6 * ((float)2 / 3 - tg);
                        }
                        else
                        {
                            ng = p;
                        }
                    }
                }

                if (tb < (float)1 / 6)
                {
                    nb = p + (q - p) * 6 * tb;
                }
                else
                {
                    if (tb < (float)1 / 2)
                    {
                        nb = q;
                    }
                    else
                    {
                        if (tb < (float)2 / 3)
                        {
                            nb = p + (q - p) * 6 * ((float)2 / 3 - tb);
                        }
                        else
                        {
                            nb = p;
                        }
                    }
                }
            }
            int r = (int)(nr * 255);
            int g = (int)(ng * 255);
            int b = (int)(nb * 255);
            int rgb = r + g * 256 + b * 256 * 256;
            return rgb;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            Table_Color.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button248.Enabled = true;
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
                    if (checkBox4.Checked)
                    {
                        int r1 = panel1.BackColor.R;
                        int g1 = panel1.BackColor.G;
                        int b1 = panel1.BackColor.B;

                        int hsl1 = Rgb2Hsl(r1, g1, b1);
                        int h1 = hsl1 % 256;
                        int s1 = (hsl1 / 256) % 256;

                        int rgb2 = Hsl2Rgb(h1, s1, 250);
                        int rgb3 = Hsl2Rgb(h1, s1, 242);

                        int r2 = rgb2 % 256;
                        int g2 = (rgb2 / 256) % 256;
                        int b2 = (rgb2 / 256 / 256) % 256;

                        int r3 = rgb3 % 256;
                        int g3 = (rgb3 / 256) % 256;
                        int b3 = (rgb3 / 256 / 256) % 256;

                        panel2.BackColor = Color.FromArgb(255, r2, g2, b2);
                        panel3.BackColor = Color.FromArgb(255, r3, g3, b3);

                    }
                }
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

        private void timer1_Tick(object sender, EventArgs e)
        {
            Point p = new Point(MousePosition.X, MousePosition.Y);
            Bitmap shotImage = new Bitmap(1, 1);
            Graphics dc = Graphics.FromImage(shotImage);
            dc.CopyFromScreen(p, new Point(0, 0), shotImage.Size);
            Color c = shotImage.GetPixel(0, 0);
            panel1.BackColor = c;
            if (checkBox4.Checked)
            {
                int r1 = panel1.BackColor.R;
                int g1 = panel1.BackColor.G;
                int b1 = panel1.BackColor.B;

                int hsl1 = Rgb2Hsl(r1, g1, b1);
                int h1 = hsl1 % 256;
                int s1 = (hsl1 / 256) % 256;

                int rgb2 = Hsl2Rgb(h1, s1, 250);
                int rgb3 = Hsl2Rgb(h1, s1, 242);

                int r2 = rgb2 % 256;
                int g2 = (rgb2 / 256) % 256;
                int b2 = (rgb2 / 256 / 256) % 256;

                int r3 = rgb3 % 256;
                int g3 = (rgb3 / 256) % 256;
                int b3 = (rgb3 / 256 / 256) % 256;

                panel2.BackColor = Color.FromArgb(255, r2, g2, b2);
                panel3.BackColor = Color.FromArgb(255, r3, g3, b3);

            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            Point p = new Point(MousePosition.X, MousePosition.Y);
            Bitmap shotImage = new Bitmap(1, 1);
            Graphics dc = Graphics.FromImage(shotImage);
            dc.CopyFromScreen(p, new Point(0, 0), shotImage.Size);
            Color c = shotImage.GetPixel(0, 0);
            panel2.BackColor = c;
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            Point p = new Point(MousePosition.X, MousePosition.Y);
            Bitmap shotImage = new Bitmap(1, 1);
            Graphics dc = Graphics.FromImage(shotImage);
            dc.CopyFromScreen(p, new Point(0, 0), shotImage.Size);
            Color c = shotImage.GetPixel(0, 0);
            panel3.BackColor = c;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                panel1.Enabled = true;
                textBox1.Enabled = true;
            }
            else
            {
                panel1.Enabled = false;
                textBox1.Enabled = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                panel2.Enabled = true;
                textBox2.Enabled = true;
            }
            else
            {
                panel2.Enabled = false;
                textBox2.Enabled = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                panel3.Enabled = true;
                textBox3.Enabled = true;
            }
            else
            {
                panel3.Enabled = false;
                textBox3.Enabled = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            float tr1 = 0;
            float tr2 = 0;
            float tr3 = 0;
            try
            {
                tr1 = float.Parse(textBox1.Text.Trim()) / 100f;
                tr2 = float.Parse(textBox2.Text.Trim()) / 100f;
                tr3 = float.Parse(textBox3.Text.Trim()) / 100f;
            }
            catch{}

            int color1 = panel1.BackColor.R + panel1.BackColor.G * 256 + panel1.BackColor.B * 256 * 256;
            int color2 = panel2.BackColor.R + panel2.BackColor.G * 256 + panel2.BackColor.B * 256 * 256;
            int color3 = panel3.BackColor.R + panel3.BackColor.G * 256 + panel3.BackColor.B * 256 * 256;

            int n = 0;
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                for (int i = 1; i <= range.Count; i++)
                {
                    if (range[i].Type == Office.MsoShapeType.msoTable)
                    {
                        n += 1;
                        if (checkBox1.Checked)
                        { 
                            int cellcnt = range[i].Table.Rows[1].Cells.Count;
                            for (int j = 1; j <= cellcnt; j++)
                            {
                                range[i].Table.Rows[1].Cells[j].Shape.Fill.ForeColor.RGB = color1;
                                range[i].Table.Rows[1].Cells[j].Shape.Fill.Transparency = tr1;
                            }
                        }

                        if (checkBox2.Checked)
                        {
                            int rowcnt = range[i].Table.Rows.Count;
                            int cellcnt = range[i].Table.Rows[1].Cells.Count;
                            int beginrow = 1;
                            if (checkBox1.Checked)
                            {
                                beginrow = 2;
                            }
                            else
                            {
                                beginrow = 1;
                            }
                            if (rowcnt >= 2)
                            {
                                for (int j = beginrow; j <= rowcnt; j++)
                                {
                                    if (j % 2 == 0)
                                    {
                                        for (int k = 1; k <= cellcnt; k++)
                                        {
                                            range[i].Table.Rows[j].Cells[k].Shape.Fill.ForeColor.RGB = color2;
                                            range[i].Table.Rows[j].Cells[k].Shape.Fill.Transparency = tr2;
                                        }
                                    }
                                }
                            }
                        }

                        if (checkBox3.Checked)
                        {
                            int rowcnt = range[i].Table.Rows.Count;
                            int cellcnt = range[i].Table.Rows[1].Cells.Count;
                            int beginrow = 1;
                            if (checkBox1.Checked)
                            {
                                beginrow = 2;
                            }
                            else
                            {
                                beginrow = 1;
                            }
                            if (rowcnt >= 2)
                            {
                                for (int j = beginrow; j <= rowcnt; j++)
                                {
                                    if (j % 2 == 1)
                                    {
                                        for (int k = 1; k <= cellcnt; k++)
                                        {
                                            range[i].Table.Rows[j].Cells[k].Shape.Fill.ForeColor.RGB = color3;
                                            range[i].Table.Rows[j].Cells[k].Shape.Fill.Transparency = tr3;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (n == 0)
	            {
                    MessageBox.Show("请选中表格、单元格区域、幻灯片缩略图之一");
	            }
            }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
            {
                PowerPoint.SlideRange slides = sel.SlideRange;
                foreach (PowerPoint.Slide slide in slides)
                {
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (shape.Type == Office.MsoShapeType.msoTable)
                        {
                            if (checkBox1.Checked)
                            {
                                int cellcnt = shape.Table.Rows[1].Cells.Count;
                                for (int j = 1; j <= cellcnt; j++)
                                {
                                    shape.Table.Rows[1].Cells[j].Shape.Fill.ForeColor.RGB = color1;
                                    shape.Table.Rows[1].Cells[j].Shape.Fill.Transparency = tr1;
                                }
                            }

                            if (checkBox2.Checked)
                            {
                                int rowcnt = shape.Table.Rows.Count;
                                int cellcnt = shape.Table.Rows[1].Cells.Count;
                                int beginrow = 1;
                                if (checkBox1.Checked)
                                {
                                    beginrow = 2;
                                }
                                else
                                {
                                    beginrow = 1;
                                }
                                if (rowcnt >= 2)
                                {
                                    for (int j = beginrow; j <= rowcnt; j++)
                                    {
                                        if (j % 2 == 0)
                                        {
                                            for (int k = 1; k <= cellcnt; k++)
                                            {
                                                shape.Table.Rows[j].Cells[k].Shape.Fill.ForeColor.RGB = color2;
                                                shape.Table.Rows[j].Cells[k].Shape.Fill.Transparency = tr2; ;
                                            }
                                        }
                                    }
                                }
                            }

                            if (checkBox3.Checked)
                            {
                                int rowcnt = shape.Table.Rows.Count;
                                int cellcnt = shape.Table.Rows[1].Cells.Count;
                                int beginrow = 1;
                                if (checkBox1.Checked)
                                {
                                    beginrow = 2;
                                }
                                else
                                {
                                    beginrow = 1;
                                }
                                if (rowcnt >= 2)
                                {
                                    for (int j = beginrow; j <= rowcnt; j++)
                                    {
                                        if (j % 2 == 1)
                                        {
                                            for (int k = 1; k <= cellcnt; k++)
                                            {
                                                shape.Table.Rows[j].Cells[k].Shape.Fill.ForeColor.RGB = color3;
                                                shape.Table.Rows[j].Cells[k].Shape.Fill.Transparency = tr3;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Type == Office.MsoShapeType.msoTable)
                    {
                        n = 1;
                        if (checkBox1.Checked)
                        {
                            int cellcnt = shape.Table.Rows[1].Cells.Count;
                            for (int j = 1; j <= cellcnt; j++)
                            {
                                shape.Table.Rows[1].Cells[j].Shape.Fill.ForeColor.RGB = color1;
                                shape.Table.Rows[1].Cells[j].Shape.Fill.Transparency = tr1;
                            }
                        }

                        if (checkBox2.Checked)
                        {
                            int rowcnt = shape.Table.Rows.Count;
                            int cellcnt = shape.Table.Rows[1].Cells.Count;
                            int beginrow = 1;
                            if (checkBox1.Checked)
                            {
                                beginrow = 2;
                            }
                            else
                            {
                                beginrow = 1;
                            }
                            if (rowcnt >= 2)
                            {
                                for (int j = beginrow; j <= rowcnt; j++)
                                {
                                    if (j % 2 == 0)
                                    {
                                        for (int k = 1; k <= cellcnt; k++)
                                        {
                                            shape.Table.Rows[j].Cells[k].Shape.Fill.ForeColor.RGB = color2;
                                            shape.Table.Rows[j].Cells[k].Shape.Fill.Transparency = tr2;
                                        }
                                    }
                                }
                            }
                        }

                        if (checkBox3.Checked)
                        {
                            int rowcnt = shape.Table.Rows.Count;
                            int cellcnt = shape.Table.Rows[1].Cells.Count;
                            int beginrow = 1;
                            if (checkBox1.Checked)
                            {
                                beginrow = 2;
                            }
                            else
                            {
                                beginrow = 1;
                            }
                            if (rowcnt >= 2)
                            {
                                for (int j = beginrow; j <= rowcnt; j++)
                                {
                                    if (j % 2 == 1)
                                    {
                                        for (int k = 1; k <= cellcnt; k++)
                                        {
                                            shape.Table.Rows[j].Cells[k].Shape.Fill.ForeColor.RGB = color3;
                                            shape.Table.Rows[j].Cells[k].Shape.Fill.Transparency = tr3;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (n == 0)
                {
                    MessageBox.Show("请选中一个表格");
                }
            }
            else
            {
                MessageBox.Show("请选中一个表格");
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                int r1 = panel1.BackColor.R;
                int g1 = panel1.BackColor.G;
                int b1 = panel1.BackColor.B;

                int hsl1 = Rgb2Hsl(r1, g1, b1);
                int h1 = hsl1 % 256;
                int s1 = (hsl1 / 256) % 256;

                int rgb2 = Hsl2Rgb(h1, s1, 250);
                int rgb3 = Hsl2Rgb(h1, s1, 242);

                int r2 = rgb2 % 256;
                int g2 = (rgb2 / 256) % 256;
                int b2 = (rgb2 / 256/256) % 256;

                int r3 = rgb3 % 256;
                int g3 = (rgb3 / 256) % 256;
                int b3 = (rgb3 / 256 / 256) % 256;

                panel2.BackColor = Color.FromArgb(255, r2, g2, b2);
                panel3.BackColor = Color.FromArgb(255, r3, g3, b3);

            }
        }

        private void timer4_Tick(object sender, EventArgs e)
        {
            Point p = new Point(MousePosition.X, MousePosition.Y);
            Bitmap shotImage = new Bitmap(1, 1);
            Graphics dc = Graphics.FromImage(shotImage);
            dc.CopyFromScreen(p, new Point(0, 0), shotImage.Size);
            Color c = shotImage.GetPixel(0, 0);
            panel4.BackColor = c;
        }

        private void panel4_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                timer4.Enabled = true;
            }
            if (e.Button == MouseButtons.Right)
            {
                if (this.colorDialog1.ShowDialog() == DialogResult.OK)
                {
                    this.panel4.BackColor = colorDialog1.Color;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                float tr = 0;
                try
                {
                    tr = float.Parse(textBox4.Text.Trim()) / 100f;
                }
                catch { }

                int color1 = panel4.BackColor.R + panel4.BackColor.G * 256 + panel4.BackColor.B * 256 * 256;

                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                for (int i = 1; i <= range.Count; i++)
                {
                    PowerPoint.Shape shape = range[i];
                    if (shape.Type == Office.MsoShapeType.msoTable)
                    {
                        List<int> numi = new List<int>();
                        List<int> numj = new List<int>();

                        for (int k = 1; k <= shape.Table.Rows.Count; k++)
                        {
                            for (int j = 1; j <= shape.Table.Columns.Count; j++)
                            {
                                if (shape.Table.Rows[k].Cells[j].Selected)
                                {
                                    numi.Add(k);
                                    numj.Add(j);
                                }
                            }
                        }

                        numi = numi.Distinct().ToList();
                        numj = numj.Distinct().ToList();
                        int rwcnt = numi.Count();
                        int clcnt = numj.Count();

                        if (checkBox7.Checked)
                        {
                            for (int m = 0; m < rwcnt; m++)
                            {
                                for (int n = 0; n < clcnt; n++)
                                {
                                    try
                                    {
                                        float.Parse(shape.Table.Rows[numi[m]].Cells[numj[n]].Shape.TextFrame.TextRange.Text.Trim());
                                        shape.Table.Rows[numi[m]].Cells[numj[n]].Shape.Fill.ForeColor.RGB = color1;
                                        shape.Table.Rows[numi[m]].Cells[numj[n]].Shape.Fill.Transparency = tr;
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }
                            }
                        }
                        else if (!checkBox5.Checked && !checkBox6.Checked)
                        {
                            for (int m = 0; m < rwcnt; m++)
                            {
                                for (int n = 0; n < clcnt; n++)
                                {
                                    shape.Table.Rows[numi[m]].Cells[numj[n]].Shape.Fill.ForeColor.RGB = color1;
                                    shape.Table.Rows[numi[m]].Cells[numj[n]].Shape.Fill.Transparency = tr;
                                }
                            }
                        }
                        else if (checkBox5.Checked && !checkBox6.Checked)
                        {
                            for (int m = 0; m < rwcnt; m++)
                            {
                                for (int n = 0; n < clcnt; n++)
                                {
                                    if (numi[m] % 2 == numi[0] % 2)
                                    {
                                        shape.Table.Rows[numi[m]].Cells[numj[n]].Shape.Fill.ForeColor.RGB = color1;
                                        shape.Table.Rows[numi[m]].Cells[numj[n]].Shape.Fill.Transparency = tr;
                                    }
                                }
                            }
                        }
                        else if (checkBox6.Checked && !checkBox5.Checked)
                        {
                            for (int m = 0; m < clcnt; m++)
                            {
                                for (int n = 0; n < rwcnt; n++)
                                {
                                    if (numj[m] % 2 == numj[0] % 2)
                                    {
                                        shape.Table.Columns[numj[m]].Cells[numi[n]].Shape.Fill.ForeColor.RGB = color1;
                                        shape.Table.Columns[numj[m]].Cells[numi[n]].Shape.Fill.Transparency = tr;
                                    }
                                }
                            }
                        }
                        else
                        {
                            for (int m = 0; m < clcnt; m++)
                            {
                                for (int n = 0; n < rwcnt; n++)
                                {
                                    if (numj[m] % 2 == numj[0] % 2 && numi[n] % 2 == numi[0] % 2)
                                    {
                                        shape.Table.Columns[numj[m]].Cells[numi[n]].Shape.Fill.ForeColor.RGB = color1;
                                        shape.Table.Columns[numj[m]].Cells[numi[n]].Shape.Fill.Transparency = tr;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("请选中要上色的表格、单元格区域");
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked)
            {
                checkBox5.Checked = false;
                checkBox6.Checked = false;
            }
            else
            {
                checkBox5.Checked = true;
                checkBox6.Checked = false;
            }
        }

    }
}
