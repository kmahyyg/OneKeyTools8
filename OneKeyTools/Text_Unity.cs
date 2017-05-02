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
    public partial class Text_Unity : Form
    {
        public Text_Unity()
        {
            InitializeComponent();
            this.Deactivate += new EventHandler(color1_Deactivate);
        }

        void color1_Deactivate(object sender, EventArgs e)
        {
            timer1.Stop();
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

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                button2.Enabled = true;
            }
            else
            {
                button2.Enabled = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                textBox1.Enabled = true;
            }
            else
            {
                textBox1.Enabled = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                panel1.Enabled = true;
            }
            else
            {
                panel1.Enabled = false;
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                textBox2.Enabled = true;
            }
            else
            {
                textBox2.Enabled = false;
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked)
            {
                comboBox1.Enabled = true;
            }
            else
            {
                comboBox1.Enabled = false;
            }
        }

        Font font;
        private void button2_Click(object sender, EventArgs e)
        {
            if (this.fontDialog1.ShowDialog() == DialogResult.OK)
            {
                font = this.fontDialog1.Font;
                button2.Text = "已选择";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                MessageBox.Show("请选中需要统一文字的页面或文本框");
            }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
            {
                PowerPoint.SlideRange srange = sel.SlideRange;
                foreach (PowerPoint.Slide slide in srange)
                {
                    int count = slide.Shapes.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        if (slide.Shapes[i].HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            if (slide.Shapes[i].TextFrame.HasText == Office.MsoTriState.msoTrue)
                            {
                                PowerPoint.ParagraphFormat pf = slide.Shapes[i].TextFrame.TextRange.ParagraphFormat;
                                if (checkBox1.Checked)
                                {
                                    slide.Shapes[i].TextFrame.TextRange.Font.Name = font.Name;
                                    slide.Shapes[i].TextFrame.TextRange.Font.NameFarEast = font.Name;
                                    if (font.Style == FontStyle.Italic)
                                    {
                                        slide.Shapes[i].TextFrame.TextRange.Font.Italic = Office.MsoTriState.msoTrue;
                                    }
                                    if (font.Style == FontStyle.Bold)
                                    {
                                        slide.Shapes[i].TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                                    }
                                }
                                if (checkBox2.Checked)
                                {
                                    slide.Shapes[i].TextFrame.TextRange.Font.Size = float.Parse(textBox1.Text.Trim());
                                }
                                if (checkBox3.Checked)
                                {
                                    slide.Shapes[i].TextFrame.TextRange.Font.Color.RGB = panel1.BackColor.R + panel1.BackColor.G * 256 + panel1.BackColor.B * 256 * 256;
                                }
                                if (checkBox4.Checked)
                                {
                                    pf.WordWrap = Office.MsoTriState.msoTrue;
                                    pf.SpaceWithin = float.Parse(textBox2.Text.Trim());
                                }
                                if (checkBox5.Checked)
                                {
                                    if (comboBox1.Text == "两端对齐")
                                    {
                                        pf.Alignment = PowerPoint.PpParagraphAlignment.ppAlignJustify;
                                    }
                                    else if (comboBox1.Text == "左对齐")
                                    {
                                        pf.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;
                                    }
                                    else if (comboBox1.Text == "右对齐")
                                    {
                                        pf.Alignment = PowerPoint.PpParagraphAlignment.ppAlignRight;
                                    }
                                    else if (comboBox1.Text == "居中对齐")
                                    {
                                        pf.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                                    }
                                    else if (comboBox1.Text == "分散对齐")
                                    {
                                        pf.Alignment = PowerPoint.PpParagraphAlignment.ppAlignDistribute;
                                    }
                                    else if (comboBox1.Text == "顶部对齐")
                                    {
                                        slide.Shapes[i].TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop;
                                    }
                                    else if (comboBox1.Text == "中部对齐")
                                    {
                                        slide.Shapes[i].TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                                    }
                                    else if (comboBox1.Text == "底部对齐")
                                    {
                                        slide.Shapes[i].TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorBottom;
                                    }
                                }
                            }
                            if (slide.Shapes[i].TextFrame2.HasText == Office.MsoTriState.msoTrue)
                            {
                                if (checkBox1.Checked)
                                {
                                    slide.Shapes[i].TextFrame2.TextRange.Font.Name = font.Name;
                                    slide.Shapes[i].TextFrame2.TextRange.Font.NameFarEast = font.Name;
                                    if (font.Style == FontStyle.Italic)
                                    {
                                        slide.Shapes[i].TextFrame2.TextRange.Font.Italic = Office.MsoTriState.msoTrue;
                                    }
                                    if (font.Style == FontStyle.Bold)
                                    {
                                        slide.Shapes[i].TextFrame2.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                                    }
                                }
                                if (checkBox2.Checked)
                                {
                                    slide.Shapes[i].TextFrame2.TextRange.Font.Size = float.Parse(textBox1.Text.Trim());
                                }
                                if (checkBox3.Checked)
                                {
                                    slide.Shapes[i].TextFrame2.TextRange.Font.Fill.BackColor.RGB = panel1.BackColor.R + panel1.BackColor.G * 256 + panel1.BackColor.B * 256 * 256;
                                }
                                if (checkBox4.Checked)
                                {
                                    slide.Shapes[i].TextFrame2.TextRange.ParagraphFormat.WordWrap = Office.MsoTriState.msoTrue;
                                    slide.Shapes[i].TextFrame2.TextRange.ParagraphFormat.SpaceWithin = float.Parse(textBox2.Text.Trim());
                                }
                                if (checkBox5.Checked)
                                {
                                    if (comboBox1.Text == "两端对齐")
                                    {
                                        slide.Shapes[i].TextFrame2.TextRange.ParagraphFormat.Alignment = Office.MsoParagraphAlignment.msoAlignJustify;
                                    }
                                    else if (comboBox1.Text == "左对齐")
                                    {
                                        slide.Shapes[i].TextFrame2.TextRange.ParagraphFormat.Alignment = Office.MsoParagraphAlignment.msoAlignLeft;
                                    }
                                    else if (comboBox1.Text == "右对齐")
                                    {
                                        slide.Shapes[i].TextFrame2.TextRange.ParagraphFormat.Alignment = Office.MsoParagraphAlignment.msoAlignRight;
                                    }
                                    else if (comboBox1.Text == "居中对齐")
                                    {
                                        slide.Shapes[i].TextFrame2.TextRange.ParagraphFormat.Alignment = Office.MsoParagraphAlignment.msoAlignCenter;
                                    }
                                    else if (comboBox1.Text == "分散对齐")
                                    {
                                        slide.Shapes[i].TextFrame2.TextRange.ParagraphFormat.Alignment = Office.MsoParagraphAlignment.msoAlignDistribute;
                                    }
                                    else if (comboBox1.Text == "顶部对齐")
                                    {
                                        slide.Shapes[i].TextFrame2.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop;
                                    }
                                    else if (comboBox1.Text == "中部对齐")
                                    {
                                        slide.Shapes[i].TextFrame2.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                                    }
                                    else if (comboBox1.Text == "底部对齐")
                                    {
                                        slide.Shapes[i].TextFrame2.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorBottom;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                int count = range.Count;
                for (int i = 1; i <= count; i++)
                {
                    if (range[i].HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        if (range[i].TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            PowerPoint.ParagraphFormat pf = range[i].TextFrame.TextRange.ParagraphFormat;
                            if (checkBox1.Checked)
                            {
                                range[i].TextFrame.TextRange.Font.Name = font.Name;
                                range[i].TextFrame.TextRange.Font.NameFarEast = font.Name;
                                if (font.Style == FontStyle.Italic)
                                {
                                    range[i].TextFrame.TextRange.Font.Italic = Office.MsoTriState.msoTrue;
                                }
                                if (font.Style == FontStyle.Bold)
                                {
                                    range[i].TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                                }
                            }
                            if (checkBox2.Checked)
                            {
                                range[i].TextFrame.TextRange.Font.Size = float.Parse(textBox1.Text.Trim());
                            }
                            if (checkBox3.Checked)
                            {
                                range[i].TextFrame.TextRange.Font.Color.RGB = panel1.BackColor.R + panel1.BackColor.G * 256 + panel1.BackColor.B * 256 * 256;
                            }
                            if (checkBox4.Checked)
                            {
                                pf.WordWrap = Office.MsoTriState.msoTrue;
                                pf.SpaceWithin = float.Parse(textBox2.Text.Trim());
                            }
                            if (checkBox5.Checked)
                            {
                                if (comboBox1.Text == "两端对齐")
                                {
                                    pf.Alignment = PowerPoint.PpParagraphAlignment.ppAlignJustify;
                                }
                                else if (comboBox1.Text == "左对齐")
                                {
                                    pf.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;
                                }
                                else if (comboBox1.Text == "右对齐")
                                {
                                    pf.Alignment = PowerPoint.PpParagraphAlignment.ppAlignRight;
                                }
                                else if (comboBox1.Text == "居中对齐")
                                {
                                    pf.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                                }
                                else if (comboBox1.Text == "分散对齐")
                                {
                                    pf.Alignment = PowerPoint.PpParagraphAlignment.ppAlignDistribute;
                                }
                                else if (comboBox1.Text == "顶部对齐")
                                {
                                    range[i].TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop;
                                }
                                else if (comboBox1.Text == "中部对齐")
                                {
                                    range[i].TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                                }
                                else if (comboBox1.Text == "底部对齐")
                                {
                                    range[i].TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorBottom;
                                }
                            }
                        }
                        if (range[i].TextFrame2.HasText == Office.MsoTriState.msoTrue)
                        {
                            if (checkBox1.Checked)
                            {
                                range[i].TextFrame2.TextRange.Font.Name = font.Name;
                                range[i].TextFrame2.TextRange.Font.NameFarEast = font.Name;
                                if (font.Style == FontStyle.Italic)
                                {
                                    range[i].TextFrame2.TextRange.Font.Italic = Office.MsoTriState.msoTrue;
                                }
                                if (font.Style == FontStyle.Bold)
                                {
                                    range[i].TextFrame2.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                                }
                            }
                            if (checkBox2.Checked)
                            {
                                range[i].TextFrame2.TextRange.Font.Size = float.Parse(textBox1.Text.Trim());
                            }
                            if (checkBox3.Checked)
                            {
                                range[i].TextFrame2.TextRange.Font.Fill.BackColor.RGB = panel1.BackColor.R + panel1.BackColor.G * 256 + panel1.BackColor.B * 256 * 256;
                            }
                            if (checkBox4.Checked)
                            {
                                range[i].TextFrame2.TextRange.ParagraphFormat.WordWrap = Office.MsoTriState.msoTrue;
                                range[i].TextFrame2.TextRange.ParagraphFormat.SpaceWithin = float.Parse(textBox2.Text.Trim());
                            }
                            if (checkBox5.Checked)
                            {
                                if (comboBox1.Text == "两端对齐")
                                {
                                    range[i].TextFrame2.TextRange.ParagraphFormat.Alignment = Office.MsoParagraphAlignment.msoAlignJustify;
                                }
                                else if (comboBox1.Text == "左对齐")
                                {
                                    range[i].TextFrame2.TextRange.ParagraphFormat.Alignment = Office.MsoParagraphAlignment.msoAlignLeft;
                                }
                                else if (comboBox1.Text == "右对齐")
                                {
                                    range[i].TextFrame2.TextRange.ParagraphFormat.Alignment = Office.MsoParagraphAlignment.msoAlignRight;
                                }
                                else if (comboBox1.Text == "居中对齐")
                                {
                                    range[i].TextFrame2.TextRange.ParagraphFormat.Alignment = Office.MsoParagraphAlignment.msoAlignCenter;
                                }
                                else if (comboBox1.Text == "分散对齐")
                                {
                                    range[i].TextFrame2.TextRange.ParagraphFormat.Alignment = Office.MsoParagraphAlignment.msoAlignDistribute;
                                }
                                else if (comboBox1.Text == "顶部对齐")
                                {
                                    range[i].TextFrame2.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop;
                                }
                                else if (comboBox1.Text == "中部对齐")
                                {
                                    range[i].TextFrame2.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                                }
                                else if (comboBox1.Text == "底部对齐")
                                {
                                    range[i].TextFrame2.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorBottom;
                                }
                            }
                        }
                    }
                }
            }
        }

        private void panel1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                timer1.Enabled = true;
            }
            if (e.Button == MouseButtons.Right)
            {
                if (this.colorDialog1.ShowDialog() == DialogResult.OK)
                {
                    this.panel1.BackColor = colorDialog1.Color;
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Point p = new Point(MousePosition.X, MousePosition.Y);
            Bitmap shotImage = new Bitmap(1, 1);
            Graphics dc = Graphics.FromImage(shotImage);
            dc.CopyFromScreen(p, new Point(0, 0), shotImage.Size);
            Color c = shotImage.GetPixel(0, 0);
            panel1.BackColor = c;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Text_Unity.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button85.Enabled = true;
        }
    }
}
