using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using forms = System.Windows.Forms;

namespace OneKeyTools
{
    public partial class Super_Chart : Form
    {
        public Super_Chart()
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

        private void button4_Click(object sender, EventArgs e)
        {
            Super_Chart.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button157.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                MessageBox.Show("请选选中图形");
            }
            else
            {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                float n = float.Parse(textBox1.Text.Trim());
                if (!checkBox4.Checked)
                {
                    if (checkBox1.Checked)
                    {
                        foreach (PowerPoint.Shape item in range)
                        {
                            item.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                            if (item.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                item.TextFrame2.AutoSize = Microsoft.Office.Core.MsoAutoSize.msoAutoSizeNone;
                            }
                            if (n >= 0)
                            {
                                float btm = range[1].Top + range[1].Height;
                                item.Height = item.Height * n;
                                item.Top = btm - item.Height;
                            }
                            else
                            {
                                float btm = range[1].Top + range[1].Height;
                                item.Height = item.Height * -n;
                                item.Top = btm;
                            }
                        }
                    }
                    if (checkBox2.Checked)
                    {
                        foreach (PowerPoint.Shape item in range)
                        {
                            item.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                            if (item.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                item.TextFrame2.AutoSize = Microsoft.Office.Core.MsoAutoSize.msoAutoSizeNone;
                            }
                            if (n >= 0)
                            {
                                float lft = range[1].Left;
                                item.Width = item.Width * n;
                                item.Left = lft;
                            }
                            else
                            {
                                float lft = range[1].Left;
                                item.Width = item.Width * -n;
                                item.Left = lft - item.Width;
                            }

                        }
                    }
                    if (checkBox3.Checked)
                    {
                        foreach (PowerPoint.Shape item in range)
                        {
                            float a = item.Rotation;
                            if (a == 0f)
                            {
                                a = a + 360;
                            }
                            item.Rotation = a * n;
                        }
                    }
                }
                else
                {
                    if (range.Count < 2)
                    {
                        MessageBox.Show("请在至少两个矩形条中输入【纯数字数据】，调整【第一个】数据矩形的长度或宽度，然后全选数据矩形单击确定按钮。PS：仅支持矩形条的宽度和高度");
                    }
                    else
                    {
                        if (checkBox1.Checked)
	                    {
                            int nc = 0;
                            foreach (PowerPoint.Shape item in range)
                            {
                                if (item.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoFalse || item.TextFrame.TextRange.Text=="")
                                {
                                    nc = nc + 1;
                                }
                            }
                            if (nc == 0)
                            {
                                float a = range[1].Height;
                                float b = float.Parse(range[1].TextFrame.TextRange.Text.Trim());
                                for (int i = 2; i <= range.Count; i++)
                                {
                                    range[i].TextFrame2.AutoSize = Microsoft.Office.Core.MsoAutoSize.msoAutoSizeNone;
                                    float c = float.Parse(range[i].TextFrame.TextRange.Text.Trim());
                                    if (c >= 0)
                                    {
                                        float ltb = range[1].Top + range[1].Height;
                                        range[i].Height = a / b * c;
                                        range[i].Top = ltb - range[i].Height;
                                    }
                                    else
                                    {
                                        float ltb = range[1].Top + range[1].Height;
                                        range[i].Height = a / b * -c;
                                        range[i].Top = ltb;
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("有"+nc+"个形状里没有添加数据，请先添加");
                            }
	                    }
                        if (checkBox2.Checked)
                        {
                            int nc = 0;
                            foreach (PowerPoint.Shape item in range)
                            {
                                if (item.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoFalse || item.TextFrame.TextRange.Text == "")
                                {
                                    nc = nc + 1;
                                }
                            }
                            if (nc == 0)
                            {
                                float a = range[1].Width;
                                float b = float.Parse(range[1].TextFrame.TextRange.Text.Trim());
                                for (int i = 2; i <= range.Count; i++)
                                {
                                    range[i].TextFrame2.AutoSize = Microsoft.Office.Core.MsoAutoSize.msoAutoSizeNone;
                                    float c = float.Parse(range[i].TextFrame.TextRange.Text.Trim());
                                    if (c >= 0)
                                    {
                                        float lft = range[1].Left;
                                        range[i].Width = a / b * c;
                                        range[i].Left = lft;
                                    }
                                    else
                                    {
                                        float lft = range[1].Left;
                                        range[i].Width = a / b * -c;
                                        range[i].Left = lft - range[i].Width;
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("有" + nc + "个形状里没有添加数据，请先添加");
                            }
                        }
                    }
                }
            }
        }

        private void checkBox4_Click(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                checkBox3.Enabled = false;
                textBox1.Enabled = false;
            }
            else
            {
                checkBox3.Enabled = true;
                textBox1.Enabled = true;
            }
        }


    }
}
