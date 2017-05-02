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
using forms = System.Windows.Forms;

namespace OneKeyTools
{
    public partial class ThreeD_Copy : Form
    {
        public ThreeD_Copy()
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
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone || sel.ShapeRange.Count > 1)
            {
                forms.MessageBox.Show("只支持单个图形的旋转复制");
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
                PowerPoint.Shape shape = range[1];

                int n = int.Parse(textBox1.Text.Trim()) - 1;
                float r0 = float.Parse(textBox2.Text.Trim());
                for (int i = 1; i <= n; i++)
                {
                    PowerPoint.Shape nshape = shape.Duplicate()[1];
                    nshape.ThreeD.IncrementRotationX(-r0 * i);
                    if (checkBox1.Checked)
                    {
                        float a = shape.ThreeD.BevelTopDepth;
                        float b = shape.ThreeD.BevelBottomDepth;
                        float c = shape.ThreeD.Depth;
                        if (shape.ThreeD.Visible != Office.MsoTriState.msoTrue)
                        {
                            shape.ThreeD.Visible = Office.MsoTriState.msoTrue;
                            c = shape.ThreeD.Depth - 36;
                        }
                        shape.ThreeD.Z = (float)(shape.Width / 2 / Math.Tan(r0 * Math.PI / 360)) + a + b + c;
                        nshape.ThreeD.Z = shape.ThreeD.Z;
                    }
                    nshape.Left = shape.Left;
                    nshape.Top = shape.Top;
                    nshape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone || sel.ShapeRange.Count > 1)
            {
                forms.MessageBox.Show("只支持单个图形的旋转复制");
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
                PowerPoint.Shape shape = range[1];

                int n = int.Parse(textBox1.Text.Trim()) - 1;
                float r0 = float.Parse(textBox2.Text.Trim());
                for (int i = 1; i <= n; i++)
                {
                    PowerPoint.Shape nshape = shape.Duplicate()[1];
                    nshape.ThreeD.IncrementRotationY(-r0 * i);
                    if (checkBox1.Checked)
                    {
                        float a = shape.ThreeD.BevelTopDepth;
                        float b = shape.ThreeD.BevelBottomDepth;
                        float c = shape.ThreeD.Depth;
                        if (shape.ThreeD.Visible != Office.MsoTriState.msoTrue)
                        {
                            shape.ThreeD.Visible = Office.MsoTriState.msoTrue;
                            c = shape.ThreeD.Depth - 36;
                        }
                        shape.ThreeD.Z = (float)(shape.Width / 2 / Math.Tan(r0 * Math.PI / 360)) + a + b + c;
                        nshape.ThreeD.Z = shape.ThreeD.Z;
                    }
                    nshape.Left = shape.Left;
                    nshape.Top = shape.Top;
                    nshape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone || sel.ShapeRange.Count > 1)
            {
                forms.MessageBox.Show("只支持单个图形的旋转复制");
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
                PowerPoint.Shape shape = range[1];

                int n = int.Parse(textBox1.Text.Trim()) - 1;
                float r0 = float.Parse(textBox2.Text.Trim());
                for (int i = 1; i <= n; i++)
                {
                    PowerPoint.Shape nshape = shape.Duplicate()[1];
                    nshape.ThreeD.IncrementRotationZ(-r0 * i);
                    if (checkBox1.Checked)
                    {
                        float a = shape.ThreeD.BevelTopDepth;
                        float b = shape.ThreeD.BevelBottomDepth;
                        float c = shape.ThreeD.Depth;
                        if (shape.ThreeD.Visible != Office.MsoTriState.msoTrue)
                        {
                            shape.ThreeD.Visible = Office.MsoTriState.msoTrue;
                            c = shape.ThreeD.Depth - 36;
                        }
                        shape.ThreeD.Z = (float)(shape.Width / 2 / Math.Tan(r0 * Math.PI / 360)) + a + b + c;
                        nshape.ThreeD.Z = shape.ThreeD.Z;
                    }
                    nshape.Left = shape.Left;
                    nshape.Top = shape.Top;
                    nshape.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                }
            }
        }

        private void ThreeDcopy_FormClosed(object sender, FormClosedEventArgs e)
        {
            Globals.Ribbons.Ribbon1.button57.Enabled = true;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            label5.Text = Math.Round(float.Parse(textBox2.Text.Trim()) * (float.Parse(textBox1.Text.Trim())),1)+"°";
            int n = (int)(360f / float.Parse(textBox2.Text.Trim()));
            label7.Text = "提示：围成圈需复制 " + n + " 个";
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            label5.Text = Math.Round(float.Parse(textBox2.Text.Trim()) * (float.Parse(textBox1.Text.Trim())),1) + "°";
            double n2 = Math.Round(360f / float.Parse(textBox1.Text.Trim()), 1);
            label7.Text = "提示：围成圈需设置 " + n2 + " °";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ThreeD_Copy.ActiveForm.Close();
        }

    }
}