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
    public partial class SuperLine : Form
    {
        public SuperLine()
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

        private void button13_Click(object sender, EventArgs e)
        {
            SuperLine.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button245.Enabled = true;
        }

        private double CMtoP(double cm)
        {
            double p = cm * 72 / 2.54;
            return p;
        }

        private float CMtoP1(double cm)
        {
            float p = (float)(cm * 72 / 2.54);
            return p;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text.Trim() != "")
            {
                float radio = float.Parse(textBox1.Text.Trim());
                if (radio < 0 || radio > 360)
                {
                    MessageBox.Show("模式1的角度范围为[0°,360°]；模式2的角度范围为[0°,180°]");
                    textBox1.Text = "";
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            float radio = float.Parse(textBox1.Text.Trim());
            textBox1.Text = (radio + 1).ToString();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            float radio = float.Parse(textBox1.Text.Trim());
            textBox1.Text = (radio - 1).ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                float nw = app.ActivePresentation.PageSetup.SlideWidth / 2;
                float nh = app.ActivePresentation.PageSetup.SlideHeight / 2;
                if (textBox1.Text.Trim() != "" && textBox2.Text.Trim() != "")
                {
                    double radio = double.Parse(textBox1.Text.Trim());
                    float length = float.Parse(textBox2.Text.Trim());
                    float pl = CMtoP1(length);
                    float nx = 0;
                    float ny = 0;
                    if ((radio >= 0 && radio < 90) || radio == 360)
                    {
                        float a = (float)Math.Cos(radio * Math.PI / 180);
                        nx = pl * (float)Math.Cos(radio * Math.PI / 180);
                        ny = -pl * (float)Math.Sin(radio * Math.PI / 180);
                    }
                    else if (radio >= 90 && radio < 180)
                    {
                        radio = 180 - radio;
                        nx = -pl * (float)Math.Cos(radio * Math.PI / 180);
                        ny = -pl * (float)Math.Sin(radio * Math.PI / 180);
                    }
                    else if (radio >= 180 && radio < 270)
                    {
                        radio = 270 - radio;
                        nx = -pl * (float)Math.Cos(radio * Math.PI / 180);
                        ny = pl * (float)Math.Sin(radio * Math.PI / 180);
                    }
                    else if (radio >= 270 && radio < 360)
                    {
                        radio = 360 - radio;
                        nx = pl * (float)Math.Cos(radio * Math.PI / 180);
                        ny = pl * (float)Math.Sin(radio * Math.PI / 180);
                    }
                    PowerPoint.FreeformBuilder fb = slide.Shapes.BuildFreeform(Office.MsoEditingType.msoEditingAuto, nw, nh);
                    fb.AddNodes(Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingAuto, nw + nx, nh + ny);
                    fb.ConvertToShape().Select();
                }
            }
            else
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                if (range[1].Type == Office.MsoShapeType.msoAutoShape || range[1].Type == Office.MsoShapeType.msoFreeform)
                {
                    if (range[1].Type != Office.MsoShapeType.msoFreeform)
                    {
                        range[1].Nodes.Insert(1, Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingCorner, 0f, 0f, 0f, 0f, 0f, 0f);
                        range[1].Nodes.Delete(2);
                    }
                    int nodecount = range[1].Nodes.Count;
                    float nw = range[1].Nodes[nodecount].Points[1, 1];
                    float nh = range[1].Nodes[nodecount].Points[1, 2];
                    if (textBox1.Text.Trim() != "" && textBox2.Text.Trim() != "")
                    {
                        double radio = double.Parse(textBox1.Text.Trim());
                        float length = float.Parse(textBox2.Text.Trim());
                        float pl = CMtoP1(length);
                        float nx = 0;
                        float ny = 0;
                        if ((radio >= 0 && radio < 90) || radio == 360)
                        {
                            nx = pl * (float)Math.Cos(radio * Math.PI / 180);
                            ny = -pl * (float)Math.Sin(radio * Math.PI / 180);
                        }
                        else if (radio >= 90 && radio < 180)
                        {
                            radio = 180 - radio;
                            nx = -pl * (float)Math.Cos(radio * Math.PI / 180);
                            ny = -pl * (float)Math.Sin(radio * Math.PI / 180);
                        }
                        else if (radio >= 180 && radio < 270)
                        {
                            radio = 270 - radio;
                            nx = -pl * (float)Math.Sin(radio * Math.PI / 180);
                            ny = pl * (float)Math.Cos(radio * Math.PI / 180);
                        }
                        else if (radio >= 270 && radio < 360)
                        {
                            radio = 360 - radio;
                            nx = pl * (float)Math.Cos(radio * Math.PI / 180);
                            ny = pl * (float)Math.Sin(radio * Math.PI / 180);
                        }
                        range[1].Nodes.Insert(nodecount, Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingCorner, nw + nx, nh + ny);
                    }
                }
                else
                {
                    MessageBox.Show("请选择AutoShape或FreeForm形状");
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                float nw = app.ActivePresentation.PageSetup.SlideWidth / 2;
                float nh = app.ActivePresentation.PageSetup.SlideHeight / 2;
                if (textBox1.Text.Trim() != "" && textBox2.Text.Trim() != "")
                {
                    double radio = double.Parse(textBox1.Text.Trim());
                    double length = double.Parse(textBox2.Text.Trim());
                    double pl = CMtoP(length);
                    double nx = 0;
                    double ny = 0;
                    if ((radio >= 0 && radio < 90) || radio == 360)
                    {
                        double a = Math.Cos(radio * Math.PI / 180);
                        nx = pl * Math.Cos(radio * Math.PI / 180);
                        ny = -pl * Math.Sin(radio * Math.PI / 180);
                    }
                    else if (radio >= 90 && radio < 180)
                    {
                        radio = 180 - radio;
                        nx = -pl * Math.Cos(radio * Math.PI / 180);
                        ny = -pl * Math.Sin(radio * Math.PI / 180);
                    }
                    else if (radio >= 180 && radio < 270)
                    {
                        radio = 270 - radio;
                        nx = -pl * Math.Sin(radio * Math.PI / 180);
                        ny = pl * Math.Cos(radio * Math.PI / 180);
                    }
                    else if (radio >= 270 && radio < 360)
                    {
                        radio = 360 - radio;
                        nx = pl * Math.Cos(radio * Math.PI / 180);
                        ny = pl * Math.Sin(radio * Math.PI / 180);
                    }
                    PowerPoint.FreeformBuilder fb = slide.Shapes.BuildFreeform(Office.MsoEditingType.msoEditingAuto, nw, nh);
                    fb.AddNodes(Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingAuto, (float)(nw + nx), (float)(nh + ny));
                    fb.ConvertToShape().Select();
                }
            }
            else
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                if (range[1].Type == Office.MsoShapeType.msoAutoShape || range[1].Type == Office.MsoShapeType.msoFreeform)
                {
                    if (range[1].Type != Office.MsoShapeType.msoFreeform)
                    {
                        range[1].Nodes.Insert(1, Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingCorner, 0f, 0f, 0f, 0f, 0f, 0f);
                        range[1].Nodes.Delete(2);
                    }
                    int nodecount = range[1].Nodes.Count;
                    if (textBox1.Text.Trim() != "" && textBox2.Text.Trim() != "")
                    {
                        double radio = double.Parse(textBox1.Text.Trim());
                        if (radio <= 180)
                        {
                            double length = double.Parse(textBox2.Text.Trim());
                            double pl = CMtoP(length);
                            double x1 = range[1].Nodes[nodecount - 1].Points[1, 1];
                            double y1 = range[1].Nodes[nodecount - 1].Points[1, 2];
                            double x2 = range[1].Nodes[nodecount].Points[1, 1];
                            double y2 = range[1].Nodes[nodecount].Points[1, 2];
                            double a1 = 0;
                            double a2 = 0;
                            double x3 = 0;
                            double y3 = 0;
                            if (Math.Abs(x1 - x2) <= 0.01)
                            {
                                x1 = x2;
                            }
                            if (Math.Abs(y1 - y2) <= 0.01)
                            {
                                y1 = y2;
                            }
                            if (y2 - y1 > 0 && x2 - x1 > 0)
                            {
                                a1 = Math.Atan(((y2 - y1) / (x2 - x1))) / Math.PI * 180;
                                if (radio < a1)
                                {
                                    a2 = a1 - radio;
                                    if (Math.Abs(a2 - (int)a2) < 0.001)
                                    {
                                        a2 = (int)a2;
                                    }
                                    else if (Math.Abs(a2 - Math.Ceiling(a2)) < 0.001)
                                    {
                                        a2 = Math.Ceiling(a2);
                                    }
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                }
                                else
                                {
                                    a2 = radio - a1;
                                    if (Math.Abs(a2 - (int)a2) < 0.001)
                                    {
                                        a2 = (int)a2;
                                    }
                                    else if (Math.Abs(a2 - Math.Ceiling(a2)) < 0.001)
                                    {
                                        a2 = Math.Ceiling(a2);
                                    }
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                }

                            }
                            else if (y2 - y1 < 0 && x2 - x1 > 0)
                            {
                                a1 = Math.Atan(((x2 - x1) / (y1 - y2))) / Math.PI * 180;
                                if (radio < a1)
                                {
                                    a2 = a1 - radio;
                                    if (Math.Abs(a2 - (int)a2) < 0.001)
                                    {
                                        a2 = (int)a2;
                                    }
                                    else if (Math.Abs(a2 - Math.Ceiling(a2)) < 0.001)
                                    {
                                        a2 = Math.Ceiling(a2);
                                    }
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                }
                                else
                                {
                                    a2 = radio - a1;
                                    if (Math.Abs(a2 - (int)a2) < 0.001)
                                    {
                                        a2 = (int)a2;
                                    }
                                    else if (Math.Abs(a2 - Math.Ceiling(a2)) < 0.001)
                                    {
                                        a2 = Math.Ceiling(a2);
                                    }
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                }
                            }
                            else if (y2 - y1 < 0 && x2 - x1 < 0)
                            {
                                a1 = Math.Atan(((y1 - y2) / (x1 - x2))) / Math.PI * 180;
                                if (radio < a1)
                                {
                                    a2 = a1 - radio;
                                    if (Math.Abs(a2 - (int)a2) < 0.001)
                                    {
                                        a2 = (int)a2;
                                    }
                                    else if (Math.Abs(a2 - Math.Ceiling(a2)) < 0.001)
                                    {
                                        a2 = Math.Ceiling(a2);
                                    }
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                }
                                else
                                {
                                    a2 = radio - a1;
                                    if (Math.Abs(a2 - (int)a2) < 0.001)
                                    {
                                        a2 = (int)a2;
                                    }
                                    else if (Math.Abs(a2 - Math.Ceiling(a2)) < 0.001)
                                    {
                                        a2 = Math.Ceiling(a2);
                                    }
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                }
                            }
                            else if (y2 - y1 > 0 && x2 - x1 < 0)
                            {
                                a1 = Math.Atan(((x1 - x2) / (y2 - y1))) / Math.PI * 180;
                                if (radio < a1)
                                {
                                    a2 = a1 - radio;
                                    if (Math.Abs(a2 - (int)a2) < 0.001)
                                    {
                                        a2 = (int)a2;
                                    }
                                    else if (Math.Abs(a2 - Math.Ceiling(a2)) < 0.001)
                                    {
                                        a2 = Math.Ceiling(a2);
                                    }
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                }
                                else
                                {
                                    a2 = radio - a1;
                                    if (Math.Abs(a2 - (int)a2) < 0.001)
                                    {
                                        a2 = (int)a2;
                                    }
                                    else if (Math.Abs(a2 - Math.Ceiling(a2)) < 0.001)
                                    {
                                        a2 = Math.Ceiling(a2);
                                    }
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                }
                            }
                            else if (Math.Abs(x2 - x1) <= 0)
                            {
                                x2 = x1;
                                if (y1 - y2 > 0)
                                {
                                    a2 = radio;
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                }
                                else if (Math.Abs(y2 - y1) <= 0)
                                {
                                    x3 = x2 + pl;
                                    y3 = y2;
                                }
                                else if (y2 - y1 > 0)
                                {
                                    a2 = radio;
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                }
                            }
                            else if (Math.Abs(y2 - y1) <= 0)
                            {
                                y2 = y1;
                                if (x1 - x2 > 0)
                                {
                                    a2 = radio;
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                }
                                else if (x2 - x1 > 0)
                                {
                                    a2 = radio;
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                }
                            }
                            range[1].Nodes.Insert(nodecount, Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingCorner, (float)x3, (float)y3);
                        }
                        else
                        {
                            MessageBox.Show("请点击“右”按钮");
                        }

                    }
                }
                else
                {
                    MessageBox.Show("请选择AutoShape或FreeForm形状");
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                float nw = app.ActivePresentation.PageSetup.SlideWidth / 2;
                float nh = app.ActivePresentation.PageSetup.SlideHeight / 2;
                if (textBox1.Text.Trim() != "" && textBox2.Text.Trim() != "")
                {
                    double radio = double.Parse(textBox1.Text.Trim());
                    double length = double.Parse(textBox2.Text.Trim());
                    double pl = CMtoP(length);
                    double nx = 0;
                    double ny = 0;
                    if ((radio >= 0 && radio < 90) || radio == 360)
                    {
                        double a = Math.Cos(radio * Math.PI / 180);
                        nx = pl * Math.Cos(radio * Math.PI / 180);
                        ny = -pl * Math.Sin(radio * Math.PI / 180);
                    }
                    else if (radio >= 90 && radio < 180)
                    {
                        radio = 180 - radio;
                        nx = -pl * Math.Cos(radio * Math.PI / 180);
                        ny = -pl * Math.Sin(radio * Math.PI / 180);
                    }
                    else if (radio >= 180 && radio < 270)
                    {
                        radio = 270 - radio;
                        nx = -pl * Math.Sin(radio * Math.PI / 180);
                        ny = pl * Math.Cos(radio * Math.PI / 180);
                    }
                    else if (radio >= 270 && radio < 360)
                    {
                        radio = 360 - radio;
                        nx = pl * Math.Cos(radio * Math.PI / 180);
                        ny = pl * Math.Sin(radio * Math.PI / 180);
                    }
                    PowerPoint.FreeformBuilder fb = slide.Shapes.BuildFreeform(Office.MsoEditingType.msoEditingAuto, nw, nh);
                    fb.AddNodes(Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingAuto, (float)(nw - nx), (float)(nh - ny));
                    fb.ConvertToShape().Select();
                }
            }
            else
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                if (range[1].Type == Office.MsoShapeType.msoAutoShape || range[1].Type == Office.MsoShapeType.msoFreeform)
                {
                    if (range[1].Type != Office.MsoShapeType.msoFreeform)
                    {
                        range[1].Nodes.Insert(1, Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingCorner, 0f, 0f, 0f, 0f, 0f, 0f);
                        range[1].Nodes.Delete(2);
                    }
                    int nodecount = range[1].Nodes.Count;
                    if (textBox1.Text.Trim() != "" && textBox2.Text.Trim() != "")
                    {
                        double radio = 180 - double.Parse(textBox1.Text.Trim());
                        if (radio <= 180)
                        {
                            double length = double.Parse(textBox2.Text.Trim());
                            double pl = CMtoP(length);
                            double x1 = range[1].Nodes[nodecount - 1].Points[1, 1];
                            double y1 = range[1].Nodes[nodecount - 1].Points[1, 2];
                            double x2 = range[1].Nodes[nodecount].Points[1, 1];
                            double y2 = range[1].Nodes[nodecount].Points[1, 2];
                            double a1 = 0;
                            double a2 = 0;
                            double x3 = 0;
                            double y3 = 0;
                            if (Math.Abs(x1 - x2) <= 0.01)
                            {
                                x1 = x2;
                            }
                            if (Math.Abs(y1 - y2) <= 0.01)
                            {
                                y1 = y2;
                            }
                            if (y2 - y1 > 0 && x2 - x1 > 0)
                            {
                                a1 = Math.Atan(((y2 - y1) / (x2 - x1))) / Math.PI * 180;
                                if (radio < a1)
                                {
                                    a2 = a1 - radio;
                                    if (Math.Abs(a2 - (int)a2) < 0.001)
                                    {
                                        a2 = (int)a2;
                                    }
                                    else if (Math.Abs(a2 - Math.Ceiling(a2)) < 0.001)
                                    {
                                        a2 = Math.Ceiling(a2);
                                    }
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                }
                                else
                                {
                                    a2 = radio - a1;
                                    if (Math.Abs(a2 - (int)a2) < 0.001)
                                    {
                                        a2 = (int)a2;
                                    }
                                    else if (Math.Abs(a2 - Math.Ceiling(a2)) < 0.001)
                                    {
                                        a2 = Math.Ceiling(a2);
                                    }
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                }

                            }
                            else if (y2 - y1 < 0 && x2 - x1 > 0)
                            {
                                a1 = Math.Atan(((x2 - x1) / (y1 - y2))) / Math.PI * 180;
                                if (radio < a1)
                                {
                                    a2 = a1 - radio;
                                    if (Math.Abs(a2 - (int)a2) < 0.001)
                                    {
                                        a2 = (int)a2;
                                    }
                                    else if (Math.Abs(a2 - Math.Ceiling(a2)) < 0.001)
                                    {
                                        a2 = Math.Ceiling(a2);
                                    }
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                }
                                else
                                {
                                    a2 = radio - a1;
                                    if (Math.Abs(a2 - (int)a2) < 0.001)
                                    {
                                        a2 = (int)a2;
                                    }
                                    else if (Math.Abs(a2 - Math.Ceiling(a2)) < 0.001)
                                    {
                                        a2 = Math.Ceiling(a2);
                                    }
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                }
                            }
                            else if (y2 - y1 < 0 && x2 - x1 < 0)
                            {
                                a1 = Math.Atan(((y1 - y2) / (x1 - x2))) / Math.PI * 180;
                                if (radio < a1)
                                {
                                    a2 = a1 - radio;
                                    if (Math.Abs(a2 - (int)a2) < 0.001)
                                    {
                                        a2 = (int)a2;
                                    }
                                    else if (Math.Abs(a2 - Math.Ceiling(a2)) < 0.001)
                                    {
                                        a2 = Math.Ceiling(a2);
                                    }
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                }
                                else
                                {
                                    a2 = radio - a1;
                                    if (Math.Abs(a2 - (int)a2) < 0.001)
                                    {
                                        a2 = (int)a2;
                                    }
                                    else if (Math.Abs(a2 - Math.Ceiling(a2)) < 0.001)
                                    {
                                        a2 = Math.Ceiling(a2);
                                    }
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                }
                            }
                            else if (y2 - y1 > 0 && x2 - x1 < 0)
                            {
                                a1 = Math.Atan(((x1 - x2) / (y2 - y1))) / Math.PI * 180;
                                if (radio < a1)
                                {
                                    a2 = a1 - radio;
                                    if (Math.Abs(a2 - (int)a2) < 0.001)
                                    {
                                        a2 = (int)a2;
                                    }
                                    else if (Math.Abs(a2 - Math.Ceiling(a2)) < 0.001)
                                    {
                                        a2 = Math.Ceiling(a2);
                                    }
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                }
                                else
                                {
                                    a2 = radio - a1;
                                    if (Math.Abs(a2 - (int)a2) < 0.001)
                                    {
                                        a2 = (int)a2;
                                    }
                                    else if (Math.Abs(a2 - Math.Ceiling(a2)) < 0.001)
                                    {
                                        a2 = Math.Ceiling(a2);
                                    }
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                }
                            }
                            else if (Math.Abs(x2 - x1) <= 0)
                            {
                                x2 = x1;
                                if (y1 - y2 > 0)
                                {
                                    a2 = radio;
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                }
                                else if (Math.Abs(y2 - y1) <= 0)
                                {
                                    x3 = x2 - pl;
                                    y3 = y2;
                                }
                                else if (y2 - y1 > 0)
                                {
                                    a2 = radio;
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                    }
                                }
                            }
                            else if (Math.Abs(y2 - y1) <= 0)
                            {
                                y2 = y1;
                                if (x1 - x2 > 0)
                                {
                                    a2 = radio;
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 + Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                }
                                else if (x2 - x1 > 0)
                                {
                                    a2 = radio;
                                    if (a2 >= 0 && a2 < 90)
                                    {
                                        x3 = x2 + Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                    else if (a2 >= 90 && a2 <= 180)
                                    {
                                        a2 = 180 - a2;
                                        x3 = x2 - Math.Cos(a2 * Math.PI / 180) * pl;
                                        y3 = y2 - Math.Sin(a2 * Math.PI / 180) * pl;
                                    }
                                }
                            }
                            range[1].Nodes.Insert(nodecount, Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingCorner, (float)x3, (float)y3);
                        }
                        else
                        {
                            MessageBox.Show("请点击“左”按钮");
                        }

                    }
                }
                else
                {
                    MessageBox.Show("请选择AutoShape或FreeForm形状");
                }
            }
        }

    }
}
