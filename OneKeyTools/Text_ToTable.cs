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
    public partial class Text_ToTable : Form
    {
        public Text_ToTable()
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

        private void button3_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides || sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                MessageBox.Show("请选中文本框");
            }
            else
            {
                try
                {
                    int num = int.Parse(textBox1.Text.Trim());
                    if (num >= 1)
                    {
                        PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                        float pw = app.ActivePresentation.PageSetup.SlideWidth;
                        float ph = app.ActivePresentation.PageSetup.SlideHeight;
                        if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                        {
                            PowerPoint.ShapeRange range = sel.ShapeRange;
                            if (sel.HasChildShapeRange)
                            {
                                range = sel.ChildShapeRange;
                            }
                            if (range.Count == 1)
                            {
                                if (range[1].TextEffect.Text != "")
                                {
                                    String[] arr = range[1].TextEffect.Text.Trim().Split(char.Parse(" "), char.Parse("\v"), char.Parse("\r")).ToArray();
                                    int txtcnt = arr.Count();
                                    if (radioButton1.Checked)
                                    {
                                        PowerPoint.Shape table = slide.Shapes.AddTable(num, (int)Math.Ceiling((float)txtcnt / (float)num), pw / 4, ph / 4, pw / 2, ph / 2);
                                        table.Table.FirstRow = false;
                                        for (int i = 1; i <= txtcnt; i++)
                                        {
                                            table.Table.Columns[(i - 1) / num + 1].Cells[(i - 1) % num + 1].Shape.TextFrame.TextRange.Text = arr[i - 1];
                                        }
                                    }
                                    else if (radioButton2.Checked)
                                    {
                                        PowerPoint.Shape table = slide.Shapes.AddTable((int)Math.Ceiling((float)txtcnt / (float)num), num, pw / 4, ph / 4, pw / 2, ph / 2);
                                        table.Table.FirstRow = false;
                                        for (int i = 1; i <= txtcnt; i++)
                                        {
                                            table.Table.Rows[(i - 1) / num + 1].Cells[(i - 1) % num + 1].Shape.TextFrame.TextRange.Text = arr[i - 1];
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("无文本");
                                }
                            }
                            else
                            {
                                List<string> txts = new List<string>();
                                for (int i = 1; i <= range.Count; i++)
                                {
                                    PowerPoint.Shape txt = range[i];
                                    if (txt.TextEffect.Text != "")
                                    {
                                        txts.Add(txt.TextEffect.Text);
                                    }
                                }
                                int txtcnt = txts.Count();
                                if (radioButton1.Checked)
                                {
                                    PowerPoint.Shape table = slide.Shapes.AddTable(num, (int)Math.Ceiling((float)txtcnt / (float)num), pw / 4, ph / 4, pw / 2, ph / 2);
                                    table.Table.FirstRow = false;
                                    for (int i = 1; i <= txtcnt; i++)
                                    {
                                        table.Table.Columns[(i - 1) / num + 1].Cells[(i - 1) % num + 1].Shape.TextFrame.TextRange.Text = txts[i - 1];
                                    }
                                }
                                else if (radioButton2.Checked)
                                {
                                    PowerPoint.Shape table = slide.Shapes.AddTable((int)Math.Ceiling((float)txtcnt / (float)num), num, pw / 4, ph / 4, pw / 2, ph / 2);
                                    table.Table.FirstRow = false;
                                    for (int i = 1; i <= txtcnt; i++)
                                    {
                                        table.Table.Rows[(i - 1) / num + 1].Cells[(i - 1) % num + 1].Shape.TextFrame.TextRange.Text = txts[i - 1];
                                    }
                                }
                            }
                        }
                        else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                        {
                            string txt = sel.TextRange.Text;
                            String[] arr = txt.Trim().Split(char.Parse(" "),char.Parse("\v"), char.Parse("\r")).ToArray();
                            int txtcnt = arr.Count();
                            if (radioButton1.Checked)
                            {
                                PowerPoint.Shape table = slide.Shapes.AddTable(num, (int)Math.Ceiling((float)txtcnt / (float)num), pw / 4, ph / 4, pw / 2, ph / 2);
                                table.Table.FirstRow = false;
                                for (int i = 1; i <= txtcnt; i++)
                                {
                                    table.Table.Columns[(i - 1) / num + 1].Cells[(i - 1) % num + 1].Shape.TextFrame.TextRange.Text = arr[i - 1];
                                }
                            }
                            else if (radioButton2.Checked)
                            {
                                PowerPoint.Shape table = slide.Shapes.AddTable((int)Math.Ceiling((float)txtcnt / (float)num), num, pw / 4, ph / 4, pw / 2, ph / 2);
                                table.Table.FirstRow = false;
                                for (int i = 1; i <= txtcnt; i++)
                                {
                                    table.Table.Rows[(i - 1) / num + 1].Cells[(i - 1) % num + 1].Shape.TextFrame.TextRange.Text = arr[i - 1];
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("请选中文本框或文本");
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("数值输入有误，请重新输入");
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Text_ToTable.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.splitButton9.Enabled = true;
        }
    }
}
