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
    public partial class Copy_Rectangle : Form
    {
        public Copy_Rectangle()
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
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                MessageBox.Show("请先选中形状或单字符文本框");
            }
            else
            {
                int row = int.Parse(textBox1.Text.Trim());
                int column = int.Parse(textBox2.Text.Trim());

                if (radioButton1.Checked)
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
                    for (int k = 1; k <= range.Count; k++)
                    {
                        PowerPoint.Shape shape = range[k];
                        for (int i = 1; i <= row; i++)
                        {
                            for (int j = 1; j <= column; j++)
                            {
                                PowerPoint.Shape nshape = shape.Duplicate()[1];
                                nshape.Top = shape.Top + shape.Height * (i - 1);
                                nshape.Left = shape.Left + shape.Width * j;
                            }
                        }
                    }

                }
                if (radioButton2.Checked)
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
                    for (int k = 1; k <= range.Count; k++)
                    {
                        PowerPoint.Shape shape = range[k];
                        string txt = shape.TextFrame.TextRange.Text;
                        PowerPoint.Shape nshape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, shape.Left + shape.Width, shape.Top, shape.Width, shape.Height);
                        int count = row * column;
                        string result = txt.PadLeft(count, '$').Replace("$", txt);
                        for (int i = 1; i < row; i++)
                        {
                            result = result.Insert(txt.Length * column * i + 2 * (i - 1), Environment.NewLine);
                        }
                        nshape.TextFrame.TextRange.Text = result;
                    }    
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Copy_Rectangle.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button29.Enabled = true;
        }
    }
}
