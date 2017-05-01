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
    public partial class Points_Adjust : Form
    {
        public Points_Adjust()
        {
            InitializeComponent();
        }

        PowerPoint.Application app = Globals.ThisAddIn.Application;

        private void button1_Click(object sender, EventArgs e)
        {
            Points_Adjust.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button177.Enabled = true;
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

        private void button2_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("请先选中一个矢量形状");
            }
            else
            {
                PowerPoint.Shape shape = sel.ShapeRange[1];
                if (sel.HasChildShapeRange)
                {
                    shape = sel.ChildShapeRange[1];
                }
                if (shape.Type == Office.MsoShapeType.msoAutoShape || shape.Type == Office.MsoShapeType.msoFreeform)
                {
                    if (shape.Type == Office.MsoShapeType.msoAutoShape)
                    {
                        shape.Nodes.Insert(1, Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingCorner, 0f, 0f, 0f, 0f, 0f, 0f);
                        shape.Nodes.Delete(2);
                    }
                    int count = shape.Nodes.Count;
                    comboBox1.Items.Clear();
                    for (int i = 1; i <= count; i++)
                    {
                        comboBox1.Items.Add(i);
                    }
                }
                else
                {
                    MessageBox.Show("所选图形不可编辑顶点");
                }
                
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.Trim() == "")
            {
                MessageBox.Show("请先选择顶点序号");
            }
            else
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
                    for (int i = 1; i <= count; i++)
                    {
                        PowerPoint.Shape shape = range[i];
                        if (shape.Type == Office.MsoShapeType.msoAutoShape)
                        {
                            shape.Nodes.Insert(1, Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingCorner, 0f, 0f, 0f, 0f, 0f, 0f);
                            shape.Nodes.Delete(2);
                        }
                        else if (shape.Type == Office.MsoShapeType.msoFreeform)
                        {
                            int np = int.Parse(comboBox1.Text.Trim());
                            if (np <= shape.Nodes.Count)
                            {
                                float x = shape.Nodes[np].Points[1, 1];
                                float y = shape.Nodes[np].Points[1, 2];
                                float n = float.Parse(textBox1.Text.Trim()) * 72 / 2.54f;
                                shape.Nodes.SetPosition(np, x, y - n);
                            }
                        }
                    }
                } 
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.Trim() == "")
            {
                MessageBox.Show("请先选择顶点序号");
            }
            else
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
                    for (int i = 1; i <= count; i++)
                    {
                        PowerPoint.Shape shape = range[i];
                        if (shape.Type == Office.MsoShapeType.msoAutoShape)
                        {
                            shape.Nodes.Insert(1, Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingCorner, 0f, 0f, 0f, 0f, 0f, 0f);
                            shape.Nodes.Delete(2);
                        }
                        else if (shape.Type == Office.MsoShapeType.msoFreeform)
                        {
                            int np = int.Parse(comboBox1.Text.Trim());
                            if (np <= shape.Nodes.Count)
                            {
                                float x = shape.Nodes[np].Points[1, 1];
                                float y = shape.Nodes[np].Points[1, 2];
                                float n = float.Parse(textBox1.Text.Trim()) * 72 / 2.54f;
                                shape.Nodes.SetPosition(np, x + n, y);
                            }
                        }
                    }
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.Trim() == "")
            {
                MessageBox.Show("请先选择顶点序号");
            }
            else
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
                    for (int i = 1; i <= count; i++)
                    {
                        PowerPoint.Shape shape = range[i];
                        if (shape.Type == Office.MsoShapeType.msoAutoShape)
                        {
                            shape.Nodes.Insert(1, Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingCorner, 0f, 0f, 0f, 0f, 0f, 0f);
                            shape.Nodes.Delete(2);
                        }
                        else if (shape.Type == Office.MsoShapeType.msoFreeform)
                        {
                            int np = int.Parse(comboBox1.Text.Trim());
                            if (np <= shape.Nodes.Count)
                            {
                                float x = shape.Nodes[np].Points[1, 1];
                                float y = shape.Nodes[np].Points[1, 2];
                                float n = float.Parse(textBox1.Text.Trim()) * 72 / 2.54f;
                                shape.Nodes.SetPosition(np, x, y + n);
                            }
                        }
                    }
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.Trim() == "")
            {
                MessageBox.Show("请先选择顶点序号");
            }
            else
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
                    for (int i = 1; i <= count; i++)
                    {
                        PowerPoint.Shape shape = range[i];
                        if (shape.Type == Office.MsoShapeType.msoAutoShape)
                        {
                            shape.Nodes.Insert(1, Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingCorner, 0f, 0f, 0f, 0f, 0f, 0f);
                            shape.Nodes.Delete(2);
                        }
                        else if (shape.Type == Office.MsoShapeType.msoFreeform)
                        {
                            int np = int.Parse(comboBox1.Text.Trim());
                            if (np <= shape.Nodes.Count)
                            {
                                float x = shape.Nodes[np].Points[1, 1];
                                float y = shape.Nodes[np].Points[1, 2];
                                float n = float.Parse(textBox1.Text.Trim()) * 72 / 2.54f;
                                shape.Nodes.SetPosition(np, x - n, y);
                            }
                        }
                    }
                }
            } 
        }

        int no = 0;
        private void button7_Click(object sender, EventArgs e)
        {
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            if (no == 1)
            {
                foreach (PowerPoint.Shape item in slide.Shapes)
                {
                    if (item.Name == "toval")
                    {
                        item.Delete();
                        no = 0;
                    }
                }
            }
            else
            {
                string np = comboBox1.Text.Trim();
                if (np == "")
                {
                    MessageBox.Show("请先选择顶点序号");
                }
                else
                {
                    int np2 = int.Parse(np);
                    PowerPoint.Selection sel = app.ActiveWindow.Selection;
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
	                {
		                PowerPoint.Shape shape = sel.ShapeRange[1];
                        if (sel.HasChildShapeRange)
                        {
                            shape = sel.ChildShapeRange[1];
                        }
                        if (shape.Type == Office.MsoShapeType.msoAutoShape)
                        {
                            shape.Nodes.Insert(1, Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingCorner, 0f, 0f, 0f, 0f, 0f, 0f);
                            shape.Nodes.Delete(2);
                        }
                        else if (shape.Type == Office.MsoShapeType.msoFreeform)
                        {
                            float x = shape.Nodes[np2].Points[1, 1];
                            float y = shape.Nodes[np2].Points[1, 2];
                            PowerPoint.Shape oval = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, x - 8, y - 8, 16, 16);
                            oval.Fill.Transparency = 1f;
                            oval.Line.Weight = 0.1F;
                            oval.Line.ForeColor.RGB = 255;
                            oval.Name = "toval";
                            no = 1;
                        }
                        else
                        {
                            MessageBox.Show("请先选中一个可编辑顶点的矢量形状");
                        }
	                }
                    
                }
            }
            
        }
    }
}
