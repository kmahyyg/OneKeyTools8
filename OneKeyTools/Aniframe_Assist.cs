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
    public partial class Aniframe_Assist : Form
    {
        public Aniframe_Assist()
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
            Aniframe_Assist.ActiveForm.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            string cha = textBox1.Text.Trim();
            int cn = int.Parse(textBox2.Text.Trim());
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                PowerPoint.Shape txt = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 100, 100);
                txt.TextFrame2.TextRange.Text = new string(char.Parse(cha), cn);
                txt.Select();
            }
            else
            {
                string[] name = new string[sel.ShapeRange.Count];
                for (int i = 1; i <= sel.ShapeRange.Count; i++)
                {
                    PowerPoint.Shape txt = sel.ShapeRange[i];
                    txt.TextFrame2.TextRange.Text = txt.TextFrame2.TextRange.Text + new string(char.Parse(cha), cn);
                    name[i - 1] = txt.Name;
                }
                slide.Shapes.Range(name).Select();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                for (int j = 1; j <= sel.ShapeRange.Count; j++)
                {
                    PowerPoint.Shape txt = sel.ShapeRange[j];
                    if (txt.HasTextFrame == Office.MsoTriState.msoFalse || txt.TextFrame2.HasText == Office.MsoTriState.msoFalse)
                    {
                        MessageBox.Show("文本框中应包含字符");
                    }
                    else
                    {
                        int cn = txt.TextFrame2.TextRange.Text.Length;
                        if (txt.TextFrame2.TextRange.Text.Contains("\r") || txt.TextFrame2.TextRange.Text.Contains("\v") || txt.TextFrame2.TextRange.Text.Contains("\n"))
                        {
                            String[] arr = txt.TextFrame2.TextRange.Text.Split(char.Parse("\r"), char.Parse("\v"), char.Parse("\n")).ToArray();
                            int count = arr.Count();
                            for (int i = 0; i < count; i++)
                            {
                                if (i == 0)
                                {
                                    txt.TextFrame2.TextRange.Text = arr[i];
                                }
                                else
                                {
                                    txt.TextFrame2.TextRange.Text = txt.TextFrame2.TextRange.Text + arr[i];
                                }
                            }
                        }
                        else
                        {
                            char[] arr = txt.TextFrame2.TextRange.Text.ToCharArray();
                            for (int i = 0; i < cn; i++)
                            {
                                if (i == 0)
                                {
                                    txt.TextFrame2.TextRange.Text = arr[i] + Environment.NewLine;
                                }
                                else if (i == cn - 1)
                                {
                                    txt.TextFrame2.TextRange.Text = txt.TextFrame2.TextRange.Text + arr[i];
                                }
                                else
                                {
                                    txt.TextFrame2.TextRange.Text = txt.TextFrame2.TextRange.Text + arr[i] + Environment.NewLine;
                                }
                            }
                        }
                    } 
                }
            }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                PowerPoint.TextRange tr = sel.TextRange;
                if (tr.Text.Contains("\r") || tr.Text.Contains("\n") || tr.Text.Contains("\v"))
                {
                    for (int i = tr.Text.Length; i >= 1; i --)
                    {
                        if (tr.Characters(i).Text == "\r" || tr.Characters(i).Text == "\n" || tr.Characters(i).Text == "\v")
                        {
                            tr.Characters(i).Delete();
                        }
                    }
                }
                else
                {
                    for (int i = 1; i <= tr.Text.Length; i += 2)
                    {
                        tr.Characters(i).Text = tr.Characters(i).Text + Environment.NewLine;
                    }
                }  
            }
            else
            {
                 MessageBox.Show("请选中要分段的文本框");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                for (int i = 1; i <= sel.ShapeRange.Count; i++)
                {
                    PowerPoint.Shape txt = sel.ShapeRange[i];
                    txt.TextEffect.PresetShape = Office.MsoPresetTextEffectShape.msoTextEffectShapePlainText;
                }
            }
            else
            {
                MessageBox.Show("请选中要转换的文本框");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("请选中文本框");
            }
            else
            {
                for (int i = 1; i <= sel.ShapeRange.Count; i++)
                {
                    PowerPoint.Shape txt = sel.ShapeRange[i];
                    if (txt.TextFrame.TextRange.ParagraphFormat.SpaceWithin != 0)
                    {
                        txt.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 0;
                    }
                    else
                    {
                        txt.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1;
                    }
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                for (int i = 1; i <= sel.ShapeRange.Count; i++)
                {
                    PowerPoint.Shape txt = sel.ShapeRange[i];
                    if (txt.TextFrame.TextRange.Font.Subscript == Office.MsoTriState.msoFalse)
                    {
                        txt.TextFrame.TextRange.Font.Subscript = Office.MsoTriState.msoTrue;
                    }
                    else
                    {
                        txt.TextFrame.TextRange.Font.Subscript = Office.MsoTriState.msoFalse;
                    }
                }
            }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                if (sel.TextRange.Font.Subscript == Office.MsoTriState.msoFalse)
                {
                    sel.TextRange.Font.Subscript = Office.MsoTriState.msoTrue;
                }
                else
                {
                    sel.TextRange.Font.Subscript = Office.MsoTriState.msoFalse;
                }
            }
            else
            {
                MessageBox.Show("请选中文本框");
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = comboBox1.Text;
        }

    }
}
