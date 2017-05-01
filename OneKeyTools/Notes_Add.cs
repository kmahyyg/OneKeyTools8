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
    public partial class Notes_Add : Form
    {
        public Notes_Add()
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

        private void button2_Click(object sender, EventArgs e)
        {
            Notes_Add.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button173.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionSlides)
            {
                MessageBox.Show("请先选中要粘贴备注的幻灯片页面");
            }
            else
            {
                string txt = richTextBox1.Text.Trim();
                if (checkBox1.Checked)
                {
                    foreach (PowerPoint.Slide slide in sel.SlideRange)
                    {
                        if (checkBox2.Checked)
                        {
                            slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text = txt + Environment.NewLine + slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text;
                        }
                        else
                        {
                            slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text = slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text + Environment.NewLine + txt;
                        }    
                    }
                }
                else
                {
                    foreach (PowerPoint.Slide slide in sel.SlideRange)
                    {
                        slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text = txt;
                    }

                }
                
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                checkBox2.Enabled = true;
            }
            else
            {
                checkBox2.Enabled = false;
            }
        }
    }
}
