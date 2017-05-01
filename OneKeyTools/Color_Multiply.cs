using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OneKeyTools
{
    public partial class Color_Multiply : Form
    {
        public Color_Multiply()
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

        
        private void timer1_Tick(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && ((sel.ShapeRange[1].Type == Microsoft.Office.Core.MsoShapeType.msoGroup && sel.ShapeRange.Count == 1) || (sel.ShapeRange.Count == 2 && sel.ShapeRange[1].Type != Microsoft.Office.Core.MsoShapeType.msoGroup && sel.ShapeRange[2].Type != Microsoft.Office.Core.MsoShapeType.msoGroup)))
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
                PowerPoint.Shape shape1 = range[1];
                PowerPoint.Shape shape2 = range[2];
                if (shape1.Fill.Visible == Microsoft.Office.Core.MsoTriState.msoTrue && shape2.Fill.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    int rgb1 = shape1.Fill.ForeColor.RGB;
                    int rgb2 = shape2.Fill.ForeColor.RGB;
                    int r1 = rgb1 % 256;
                    int g1 = (rgb1 / 256) % 256;
                    int b1 = (rgb1 / 256 / 256) % 256;
                    int r2 = rgb2 % 256;
                    int g2 = (rgb2 / 256) % 256;
                    int b2 = (rgb2 / 256 / 256) % 256;
                    int nr = r1 * r2 / 255;
                    int ng = g1 * g2 / 255;
                    int nb = b1 * b2 / 255;
                    label1.Text = nr + "," + ng + "," + nb;
                    panel1.BackColor = Color.FromArgb(nr, ng, nb);
                }
                else
                {
                    label1.Text = "形状无填充";
                }
            }
            else
            {
                label1.Text = "选择2个纯色形状";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Color_Multiply.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button92.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
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
                int nr = panel1.BackColor.R;
                int ng = panel1.BackColor.G;
                int nb = panel1.BackColor.B;
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape shape = range[i];
                    shape.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                    shape.Fill.ForeColor.RGB = nr + ng * 256 + nb * 256 * 256;
                }
            }
            timer1.Start();
        }
    }
}
