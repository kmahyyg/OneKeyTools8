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
    public partial class Notes_Import : Form
    {
        public Notes_Import()
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

        private void label5_Click(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            PowerPoint.Slides slides = app.ActivePresentation.Slides;
            int count = slides.Count;
            comboBox1.Items.Clear();
            comboBox1.Text = "1";
            for (int i = 2; i <= count; i++)
            {
                comboBox1.Items.Add(i);
            }
        }

        private void note1_Load(object sender, EventArgs e)
        {
            PowerPoint.Slides slides = app.ActivePresentation.Slides;
            int count = slides.Count;
            comboBox1.Text = "1";
            for (int i = 2; i <= count; i++)
            {
                comboBox1.Items.Add(i);
            }
            textBox1.Text = Properties.Settings.Default.notesplit;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("请选中要导入到备注里的文本框，不同页的文字请先用分隔符分开");
            }
            else
            {
                PowerPoint.Slides slides = app.ActivePresentation.Slides;
                PowerPoint.Shape shape = sel.ShapeRange[1];
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextEffect.Text != "")
                {
                    string sp = textBox1.Text.Trim();
                    int n = int.Parse(comboBox1.Text.Trim());
                    string txt = shape.TextEffect.Text;
                    if (txt.Contains(sp))
                    {
                        int ts = Regex.Matches(txt,sp).Count;
                        if (ts <= slides.Count - n)
                        {
                            String[] arr = txt.Split(char.Parse(sp)).ToArray();
                            int tcount = arr.Count();
                            for (int i = 1; i <= tcount; i++)
                            {
                                 slides[n + i - 1].NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text = arr[i - 1].Trim();
                            }
                            MessageBox.Show("备注导入成功");
                        }
                        else
                        {
                            MessageBox.Show("分隔符数 > 幻灯片页数");
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

        private void button3_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.notesplit = textBox1.Text;
            Properties.Settings.Default.Save();
            MessageBox.Show("修改成功");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Notes_Import.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button167.Enabled = true;
        }


    }
}
