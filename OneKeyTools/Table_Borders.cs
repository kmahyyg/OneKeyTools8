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
    public partial class Table_Borders : Form
    {
        public Table_Borders()
        {
            InitializeComponent();
            this.Deactivate += new EventHandler(color1_Deactivate);
        }

        void color1_Deactivate(object sender, EventArgs e)
        {
            timer1.Stop();
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

        private void button1_Click(object sender, EventArgs e)
        {
            Table_Borders.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button262.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int n1 = 0;
            float tr = 0f;

            try
            {
                tr = float.Parse(textBox1.Text.Trim()) / 100f;
            }
            catch{}

            float lw = 1f;
            try
            {
                lw = (float)numericUpDown1.Value;
            }
            catch{}

            int color1 = panel1.BackColor.R + panel1.BackColor.G * 256 + panel1.BackColor.B * 256 * 256;

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

                        if (checkBox1.Checked)
                        {
                            for (int m = 0; m < rwcnt; m++)
                            {
                                for (int n = 0; n < clcnt; n++)
                                {
                                    if (shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderTop].Visible == Office.MsoTriState.msoTrue)
                                    {
                                        shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderTop].Weight = lw;
                                    }
                                    if (shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderLeft].Visible == Office.MsoTriState.msoTrue)
                                    {
                                        shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderLeft].Weight = lw;
                                    }
                                    if (shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderBottom].Visible == Office.MsoTriState.msoTrue)
                                    {
                                        shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderBottom].Weight = lw;
                                    }
                                    if (shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderRight].Visible == Office.MsoTriState.msoTrue)
                                    {
                                        shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderRight].Weight = lw;
                                    }
                                }
                            }
                        }
                        if (checkBox2.Checked)
                        {
                            for (int m = 0; m < rwcnt; m++)
                            {
                                for (int n = 0; n < clcnt; n++)
                                {
                                    if (shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderTop].Visible == Office.MsoTriState.msoTrue)
                                    {
                                        shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderTop].ForeColor.RGB = color1;
                                        shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderTop].Transparency = tr;
                                    }
                                    if (shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderLeft].Visible == Office.MsoTriState.msoTrue)
                                    {
                                        shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderLeft].ForeColor.RGB = color1;
                                        shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderLeft].Transparency = tr;
                                    }
                                    if (shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderBottom].Visible == Office.MsoTriState.msoTrue)
                                    {
                                        shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = color1;
                                        shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderBottom].Transparency = tr;
                                    }
                                    if (shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderRight].Visible == Office.MsoTriState.msoTrue)
                                    {
                                        shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderRight].ForeColor.RGB = color1;
                                        shape.Table.Rows[numi[m]].Cells[numj[n]].Borders[PowerPoint.PpBorderType.ppBorderRight].Transparency = tr;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        n1 += 1;
                    }
                }
                if (n1 != 0)
                {
                    MessageBox.Show("有 " + n1 + "个图形不是表格元素，请选中表格或单元格区域之一");
                }
            }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
            {
                PowerPoint.SlideRange slides = sel.SlideRange;
                foreach (PowerPoint.Slide slide in slides)
                {
                    foreach (PowerPoint.Shape  shape in slide.Shapes)
                    {
                        if (shape.Type == Office.MsoShapeType.msoTable)
                        {
                            if (checkBox1.Checked)
                            {
                                for (int m = 1; m <= shape.Table.Rows.Count; m++)
                                {
                                    for (int n = 1; n <= shape.Table.Columns.Count; n++)
                                    {
                                        if (shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderTop].Visible == Office.MsoTriState.msoTrue)
                                        {
                                            shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderTop].Weight = lw;
                                        }
                                        if (shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderLeft].Visible == Office.MsoTriState.msoTrue)
                                        {
                                            shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderLeft].Weight = lw;
                                        }
                                        if (shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderBottom].Visible == Office.MsoTriState.msoTrue)
                                        {
                                            shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderBottom].Weight = lw;
                                        }
                                        if (shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderRight].Visible == Office.MsoTriState.msoTrue)
                                        {
                                            shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderRight].Weight = lw;
                                        }
                                    }
                                }
                            }
                            if (checkBox2.Checked)
                            {
                                for (int m = 1; m <= shape.Table.Rows.Count; m++)
                                {
                                    for (int n = 1; n <= shape.Table.Columns.Count; n++)
                                    {
                                        if (shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderTop].Visible == Office.MsoTriState.msoTrue)
                                        {
                                            shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderTop].ForeColor.RGB = color1;
                                            shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderTop].Transparency = tr;
                                        }
                                        if (shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderLeft].Visible == Office.MsoTriState.msoTrue)
                                        {
                                            shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderLeft].ForeColor.RGB = color1;
                                            shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderLeft].Transparency = tr;
                                        }
                                        if (shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderBottom].Visible == Office.MsoTriState.msoTrue)
                                        {
                                            shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = color1;
                                            shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderBottom].Transparency = tr;
                                        }
                                        if (shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderRight].Visible == Office.MsoTriState.msoTrue)
                                        {
                                            shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderRight].ForeColor.RGB = color1;
                                            shape.Table.Rows[m].Cells[n].Borders[PowerPoint.PpBorderType.ppBorderRight].Transparency = tr;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("请选中表格、单元格区域、幻灯片缩略图之一");
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
