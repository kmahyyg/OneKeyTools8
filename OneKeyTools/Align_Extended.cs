using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneKeyTools
{
    public partial class Align_Extended : Form
    {
        public Align_Extended()
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

        private void label5_Click(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                int count = range.Count;
                int ncount = int.Parse(textBox1.Text.Trim());
                float nw = float.Parse(textBox2.Text.Trim()) * 72 / 2.54f;
                if (count >= ncount)
                {
                    for (int i = 2; i <= count; i++)
			        {
			            PowerPoint.Shape shape=range[i];
                        if (shape.LockAspectRatio == Office.MsoTriState.msoFalse)
                        {
                            shape.LockAspectRatio = Office.MsoTriState.msoTrue;
                        }
                        if (radioButton1.Checked)
                        {
                            if (i > ncount)
                            {
                                int n = (i - 1) % ncount;
                                if (n == 0)
                                {
                                    shape.Width = range[1].Width;
                                    shape.Left = range[1].Left;
                                }
                                else
                                {
                                    shape.Width = range[n + 1].Width;
                                    shape.Left = range[n].Left + range[n].Width + nw;
                                }
                                shape.Top = range[i - ncount].Top + range[i - ncount].Height + nw;
                            }
                            else
                            {
                                shape.Left = range[i - 1].Left + range[i - 1].Width + nw;
                                shape.Top = range[1].Top;
                            }
                        }
                        else if (radioButton2.Checked)
                        {
                            if (i > ncount)
                            {
                                int n = (i - 1) % ncount;
                                if (n == 0)
                                {
                                    shape.Height = range[1].Height;
                                    shape.Top = range[1].Top;
                                }
                                else
                                {
                                    shape.Height = range[n + 1].Height;
                                    shape.Top = range[n].Top + range[n].Height + nw;
                                }
                                shape.Left = range[i - ncount].Left + range[i - ncount].Width + nw;
                            }
                            else
                            {
                                shape.Left = range[1].Left;
                                shape.Top = range[i - 1].Top + range[i - 1].Height + nw;
                            }
                        }
                        
			        }
                    
                }
                else
                {
                    MessageBox.Show("多来几个图形再试试吧←_←");
                }
            }
            else
            {
                MessageBox.Show("请先选中图形");
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                float lm = app.ActivePresentation.PageSetup.SlideWidth / 2;
                float tm = app.ActivePresentation.PageSetup.SlideHeight / 2;
                float radius = float.Parse(textBox3.Text.Trim()) * 72 / 2.54f;
                float start = float.Parse(textBox4.Text.Trim());
                float end = float.Parse(textBox5.Text.Trim());
                if (range.Count > 1)
                {
                    int count = range.Count;
                    float angle = (end - start) / (count - 1);
                    for (int i = 0; i < count; i++)
                    {
                        double na = start + angle * i;
                        if (!checkBox1.Checked)
                        {
                            na = start + 360 - angle * i;
                        }
                        double nx = lm - Math.Cos(na * (float)Math.PI / 180) * radius;
                        double ny = tm - Math.Sin(na * (float)Math.PI / 180) * radius;
                        range[i + 1].Left = (float)nx - range[i + 1].Width / 2;
                        range[i + 1].Top = (float)ny - range[i + 1].Height / 2;
                    }
                }
                else
                {
                    float angle = float.Parse(textBox6.Text.Trim());
                    double na = start + angle;
                    PowerPoint.Shape nshape = range[1].Duplicate()[1];
                    if (!checkBox1.Checked)
                    {
                        na = start + 360 - angle;
                    }
                    textBox4.Text = na.ToString();
                    double nx0 = lm - Math.Cos(start * (float)Math.PI / 180) * radius;
                    double ny0 = tm - Math.Sin(start * (float)Math.PI / 180) * radius;
                    range[1].Left = (float)nx0 - range[1].Width / 2;
                    range[1].Top = (float)ny0 - range[1].Height / 2;

                    double nx = lm - Math.Cos(na * (float)Math.PI / 180) * radius;
                    double ny = tm - Math.Sin(na * (float)Math.PI / 180) * radius;
                    nshape.Left = (float)nx - nshape.Width / 2;
                    nshape.Top = (float)ny - nshape.Height / 2;
                    nshape.Select();
                } 
            }
            else
            {
                MessageBox.Show("请先选中图形");
            }
        }

        private void label8_Click(object sender, EventArgs e)
        {
            textBox4.Text = "0";
        }

        private void button13_Click(object sender, EventArgs e)
        {
            Align_Extended.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button194.Enabled = true;
        }

    }
}
