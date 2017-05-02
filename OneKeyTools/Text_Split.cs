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
using System.Text.RegularExpressions;

namespace OneKeyTools
{
    public partial class Text_Split : Form
    {
        public Text_Split()
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

        private void button1_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("请选中要拆分的文本框");
            }
            else
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.Shape shape = sel.ShapeRange[1];
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextEffect.Text != "")
                {
                    string sp = textBox1.Text.Trim();
                    string txt = shape.TextEffect.Text;
                    if (txt.Contains(sp))
                    {
                        String[] arr = txt.Split(char.Parse(sp)).ToArray();
                        int tcount = arr.Count();
                        shape.PickUp();
                        for (int i = 1; i <= tcount; i++)
                        {
                            PowerPoint.Shape text = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, shape.Left + shape.Width, shape.Top + shape.Height * (i - 1) / tcount, shape.Width, shape.Height);
                            text.TextFrame.TextRange.Text = arr[i - 1].Trim();
                            text.Apply();
                            text.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                        }
                    }
                    else
                    {
                        MessageBox.Show("找不到指定的分隔符");
                    }
                }
                else
                {
                    MessageBox.Show("所选文本框中无文字");
                }
            }
        }

        private void label5_Click(object sender, EventArgs e)
        {
            Notes_Import.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button184.Enabled = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.splittext = textBox1.Text;
            Properties.Settings.Default.Save();
            MessageBox.Show("修改成功");
        }

        private void splittext1_Load(object sender, EventArgs e)
        {
            textBox1.Text = Properties.Settings.Default.splittext;
        }
    }
}
