using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using forms = System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneKeyTools
{
    public partial class Rotation_More : Form
    {
        public Rotation_More()
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

        private void button1_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type== PowerPoint.PpSelectionType.ppSelectionNone)
            {
                MessageBox.Show("请先选中一个形状");
            }
            else
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.ShapeRange range = sel.ShapeRange;
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
                    shape.Rotation = float.Parse(textBox1.Text);
                }
            }        
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionNone)
            {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                else
                {
                    range = sel.ShapeRange;
                }
                label3.Text = range[1].Rotation+"°";
            }
            else
            {
                label3.Text = "空";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                MessageBox.Show("请先选中要复制的形状");
            }
            else
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                float rn = float.Parse(textBox2.Text);
                int scount = int.Parse(textBox3.Text);
                oshape = range[1];
                for (int i = 1; i <= scount; i++)
                {
                    PowerPoint.Shape nshape = oshape.Duplicate()[1];
                    nshape.Rotation = oshape.Rotation + i * rn;
                    nshape.Left = oshape.Left;
                    nshape.Top = oshape.Top;
                    nshape.Select();
                }
            }       
        }
        PowerPoint.Shape oshape;

        private void button3_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                forms.MessageBox.Show("请选中一个组合或两个以上形状");
            }
            else
            {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                else
                {
                    range = sel.ShapeRange;
                }
                int count = range.Count;
                if (count == 1 && range[1].Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                {
                    for (int j = 1; j <= range[1].GroupItems.Count; j++)
                    {
                        range[1].GroupItems[j].Rotation = range[1].GroupItems[1].Rotation;
                    }
                }
                else if (count > 1)
                {
                    for (int i = 1; i <= count; i++)
                    {
                        range[i].Rotation = range[1].Rotation;
                    }
                }
                else
                {
                    MessageBox.Show("请选中一个组合或两个以上形状");
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                forms.MessageBox.Show("请选中至少一个形状");
            }
            else
            {
                PowerPoint.ShapeRange range = sel.ShapeRange;
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
                    range[i].Rotation = range[i].Rotation + 180;
                }        
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                forms.MessageBox.Show("请选中至少一个形状");
            }
            else
            {
                PowerPoint.ShapeRange range = sel.ShapeRange;
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
                    range[i].Rotation = 0;
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Rotation_More.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button28.Enabled = true;
        }
    }
}
