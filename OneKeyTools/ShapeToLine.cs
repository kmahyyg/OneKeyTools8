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
    public partial class ShapeToLine : Form
    {
        public ShapeToLine()
        {
            InitializeComponent();
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

        private PowerPoint.Application app = Globals.ThisAddIn.Application;

        private void button1_Click(object sender, EventArgs e)
        {
            ShapeToLine.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button176.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.Items.Count == 0 || comboBox2.Items.Count == 0)
            {
                MessageBox.Show("请先点刷新，选择起始/终止点序号");
            }
            else
            {
                PowerPoint.Selection sel = app.ActiveWindow.Selection;
                if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("请选中形状和连接符");
                }
                else
                {
                    PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    int count = range.Count;
                    List<string> list1 = new List<string>();
                    List<string> list2 = new List<string>();
                    for (int i = 1; i <= count; i++)
                    {
                        PowerPoint.Shape shape = range[i];
                        if (shape.Type == Office.MsoShapeType.msoLine)
                        {
                            list2.Add(shape.Name);
                        }
                        else
                        {
                            list1.Add(shape.Name);
                        }
                    }
                    int count2 = list1.Count();
                    int count3 = list2.Count();
                    if (count2 == 0)
                    {
                        MessageBox.Show("请同时选中普通矢量形状（非连接符）");
                    }
                    else if (count3 == 0)
                    {
                        MessageBox.Show("请同时选中连接符");
                    }
                    else
                    {
                        int count4 = 0;
                        int fc = int.Parse(comboBox1.Text.Trim());
                        int ec = int.Parse(comboBox2.Text.Trim());
                        if (count2 < count3)
                        {
                            count4 = count2;
                        }
                        else
                        {
                            count4 = count3;
                        }
                        for (int j = 0; j < count4; j++)
                        {
                            PowerPoint.Shape shape = slide.Shapes[list2[j]];
                            if (count2 > count3)
                            {
                                shape.ConnectorFormat.BeginConnect(slide.Shapes[list1[j]], fc);
                                shape.ConnectorFormat.EndConnect(slide.Shapes[list1[j + 1]], ec);
                            }
                            else if (count2 == count3)
                            {
                                shape.ConnectorFormat.BeginConnect(slide.Shapes[list1[j]], fc);
                                if (j < count3 - 1)
                                {
                                    shape.ConnectorFormat.EndConnect(slide.Shapes[list1[j + 1]], ec);
                                }
                            }
                            else if (count2 < count3)
                            {
                                shape.ConnectorFormat.BeginConnect(slide.Shapes[list1[j]], fc);
                                if (j < count2 - 1)
                                {
                                    shape.ConnectorFormat.EndConnect(slide.Shapes[list1[j + 1]], ec);
                                }
                            }
                        }
                    }
                }
            }
            
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("请同时选中矢量形状和连接符");
            }
            else
            {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                int a = 0;
                int b = 0;
                foreach (PowerPoint.Shape item in range)
                {
                    if (a == 0)
                    {
                        if (item.Type == Office.MsoShapeType.msoAutoShape || item.Type == Office.MsoShapeType.msoFreeform)
                        {
                            a += 1;
                            b = item.ConnectionSiteCount;
                        }
                    }
                }
                comboBox1.Items.Clear();
                comboBox2.Items.Clear();
                for (int i = 1; i <= b; i++)
			    {
                    comboBox1.Items.Add(i);
                    comboBox2.Items.Add(i);
			    }
                
            }
        }
    }
}
