using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
using Office=Microsoft.Office.Core;
using forms=System.Windows.Forms;

namespace OneKeyTools
{
    public partial class Align_More : Form
    {
        public Align_More()
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
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                forms.MessageBox.Show("请至少选择两个形状");
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
                int count = range.Count;
                if (count == 1)
                {
                    if (range[1].Type == Office.MsoShapeType.msoGroup)
                    {
                        for (int i = 1; i <= range[1].GroupItems.Count; i++)
                        {
                            for (int j = 2; j <= range[1].GroupItems.Count; j++)
                            {
                                range[1].GroupItems[j].Top = range[1].GroupItems[j - 1].Top + range[1].GroupItems[j - 1].Height;
                                range[1].GroupItems[j].Left = range[1].GroupItems[j - 1].Left + range[1].GroupItems[j - 1].Width;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("请至少选择两个形状");
                    }
                }
                else
                {
                    for (int i = 2; i <= count; i++)
                    {
                        range[i].Top = range[i - 1].Top + range[i - 1].Height;
                        range[i].Left = range[i - 1].Left + range[i - 1].Width;
                    }
                } 
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                forms.MessageBox.Show("请至少选择两个形状");
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
                int count = range.Count;
                if (count == 1)
                {
                    if (range[1].Type == Office.MsoShapeType.msoGroup)
                    {
                        for (int i = 1; i <= range[1].GroupItems.Count; i++)
                        {
                            for (int j = 2; j <= range[1].GroupItems.Count; j++)
                            {
                                range[1].GroupItems[j].Top = range[1].GroupItems[j - 1].Top;
                                range[1].GroupItems[j].Left = range[1].GroupItems[j - 1].Left + range[1].GroupItems[j - 1].Width;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("请至少选择两个形状");
                    }
                }
                else
                {
                    for (int i = 2; i <= count; i++)
                    {
                        range[i].Left = range[i - 1].Left + range[i - 1].Width;
                        range[i].Top = range[i - 1].Top;
                    }
                }    
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                forms.MessageBox.Show("请至少选择两个形状");
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
                int count = range.Count;
                if (count == 1)
                {
                    if (range[1].Type == Office.MsoShapeType.msoGroup)
                    {
                        for (int i = 1; i <= range[1].GroupItems.Count; i++)
                        {
                            for (int j = 2; j <= range[1].GroupItems.Count; j++)
                            {
                                range[1].GroupItems[j].Top = range[1].GroupItems[j - 1].Top + range[1].GroupItems[j - 1].Height;
                                range[1].GroupItems[j].Left = range[1].GroupItems[j - 1].Left;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("请至少选择两个形状");
                    }
                }
                else
                {
                    for (int i = 2; i <= count; i++)
                    {
                        range[i].Top = range[i - 1].Top + range[i - 1].Height;
                        range[i].Left = range[i - 1].Left;
                    }
                }     
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                forms.MessageBox.Show("请先选中至少两个形状");
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
                int count = range.Count;
                if (count == 1)
                {
                    if (range[1].Type == Office.MsoShapeType.msoGroup)
                    {
                        for (int j = 1; j < range[1].GroupItems.Count; j++)
                        {
                            float R = range[1].GroupItems[j].Width / 2;
                            float r = range[1].GroupItems[j].Width / 2;
                            float Rl = range[1].GroupItems[j].Left;
                            float rl = range[1].GroupItems[j + 1].Left;
                            float Rt = range[1].GroupItems[j].Top;
                            float rt = range[1].GroupItems[j + 1].Top;
                            float a = rl + r - Rl - R;
                            float b = rt + r - Rt - R;
                            float c = R + r;

                            if (r + rl == R + Rl || r + rt == R + Rt)
                            {
                                if (r + rl == R + Rl && r + rt < R + Rt)
                                {
                                    range[1].GroupItems[j + 1].Top = Rt - 2 * r;
                                }

                                if (r + rl == R + Rl && r + rt > R + Rt)
                                {
                                    range[1].GroupItems[j + 1].Top = Rt + 2 * R;
                                }

                                if (r + rt == R + Rt && r + rl < R + Rl)
                                {
                                    range[1].GroupItems[j + 1].Left = Rl - 2 * r;
                                }

                                if (r + rt == R + Rt && r + rl > R + Rl)
                                {
                                    range[1].GroupItems[j + 1].Left = Rl + 2 * R;
                                }

                                if (r + rt == R + Rt && r + rl == R + Rl)
                                {
                                    range[1].GroupItems[j + 1].Left = Rl + 2 * R;
                                    range[1].GroupItems[j + 1].Top = Rt;
                                }
                            }
                            else
                            {
                                float b1 = (float)((Math.Sqrt(a * a + b * b) - c) * b / Math.Sqrt(a * a + b * b));
                                float a1 = a * b1 / b;
                                range[1].GroupItems[j + 1].Left = range[1].GroupItems[j + 1].Left - a1;
                                range[1].GroupItems[j + 1].Top = range[1].GroupItems[j + 1].Top - b1;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("请至少选择两个形状");
                    }
                }
                else
                {
                    for (int i = 1; i < count; i++)
                    {
                        float R = range[i].Width / 2;
                        float r = range[i + 1].Width / 2;
                        float Rl = range[i].Left;
                        float rl = range[i + 1].Left;
                        float Rt = range[i].Top;
                        float rt = range[i + 1].Top;
                        float a = rl + r - Rl - R;
                        float b = rt + r - Rt - R;
                        float c = R + r;

                        if (r + rl == R + Rl || r + rt == R + Rt)
                        {
                            if (r + rl == R + Rl && r + rt < R + Rt)
                            {
                                range[i + 1].Top = Rt - 2 * r;
                            }

                            if (r + rl == R + Rl && r + rt > R + Rt)
                            {
                                range[i + 1].Top = Rt + 2 * R;
                            }

                            if (r + rt == R + Rt && r + rl < R + Rl)
                            {
                                range[i + 1].Left = Rl - 2 * r;
                            }

                            if (r + rt == R + Rt && r + rl > R + Rl)
                            {
                                range[i + 1].Left = Rl + 2 * R;
                            }

                            if (r + rt == R + Rt && r + rl == R + Rl)
                            {
                                range[i + 1].Left = Rl + 2 * R;
                                range[i + 1].Top = Rt;
                            }
                        }
                        else
                        {
                            float b1 = (float)((Math.Sqrt(a * a + b * b) - c) * b / Math.Sqrt(a * a + b * b));
                            float a1 = a * b1 / b;
                            range[i + 1].Left = range[i + 1].Left - a1;
                            range[i + 1].Top = range[i + 1].Top - b1;
                        }
                    }
                } 
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                forms.MessageBox.Show("请先选中至少两个形状");
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
                int count = range.Count;
                if (count == 1)
                {
                    if (range[1].Type == Office.MsoShapeType.msoGroup)
                    {
                        for (int i = 1; i < range[1].GroupItems.Count; i++)
                        {
                            float R = range[1].GroupItems[i].Width / 2;
                            float r = range[1].GroupItems[i + 1].Width / 2;
                            float Rl = range[1].GroupItems[i].Left;
                            float rl = range[1].GroupItems[i + 1].Left;
                            float Rt = range[1].GroupItems[i].Top;
                            float rt = range[1].GroupItems[i + 1].Top;
                            float a = rl + r - Rl - R;
                            float b = rt + r - Rt - R;
                            float c = R + r;
                            if (r + rt == R + Rt)
                            {
                                if (r + rl < R + Rl)
                                {
                                    range[1].GroupItems[i + 1].Left = Rl - 2 * r;
                                }

                                if (r + rl >= R + Rl)
                                {
                                    range[1].GroupItems[i + 1].Left = Rl + 2 * R;
                                }
                            }
                            else
                            {
                                if (2 * r + rt < Rt || 2 * R + Rt < rt)
                                {
                                    forms.MessageBox.Show("超出范围，不能水平贴合");
                                }
                                else
                                {
                                    if (r + rl >= R + Rl)
                                    {
                                        float a1 = (float)Math.Sqrt(c * c - b * b);
                                        range[1].GroupItems[i + 1].Left = Rl + R + a1 - r;
                                    }
                                    if (r + rl < R + Rl)
                                    {
                                        float a1 = (float)Math.Sqrt(c * c - b * b);
                                        range[1].GroupItems[i + 1].Left = Rl + R - a1 - r;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("请至少选择两个形状");
                    }
                }
                else
                {
                    for (int i = 1; i < count; i++)
                    {
                        float R = range[i].Width / 2;
                        float r = range[i + 1].Width / 2;
                        float Rl = range[i].Left;
                        float rl = range[i + 1].Left;
                        float Rt = range[i].Top;
                        float rt = range[i + 1].Top;
                        float a = rl + r - Rl - R;
                        float b = rt + r - Rt - R;
                        float c = R + r;
                        if (r + rt == R + Rt)
                        {
                            if (r + rl < R + Rl)
                            {
                                range[i + 1].Left = Rl - 2 * r;
                            }

                            if (r + rl >= R + Rl)
                            {
                                range[i + 1].Left = Rl + 2 * R;
                            }
                        }
                        else
                        {
                            if (2 * r + rt < Rt || 2 * R + Rt < rt)
                            {
                                forms.MessageBox.Show("超出范围，不能水平贴合");
                            }
                            else
                            {
                                if (r + rl >= R + Rl)
                                {
                                    float a1 = (float)Math.Sqrt(c * c - b * b);
                                    range[i + 1].Left = Rl + R + a1 - r;
                                }
                                if (r + rl < R + Rl)
                                {
                                    float a1 = (float)Math.Sqrt(c * c - b * b);
                                    range[i + 1].Left = Rl + R - a1 - r;
                                }
                            }
                        }
                    }
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                forms.MessageBox.Show("请先选中至少两个形状");
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
                int count = range.Count;
                if (count == 1)
                {
                    if (range[1].Type == Office.MsoShapeType.msoGroup)
                    {
                        for (int i = 1; i < range[1].GroupItems.Count; i++)
                        {
                            float R = range[1].GroupItems[i].Width / 2;
                            float r = range[1].GroupItems[i + 1].Width / 2;
                            float Rl = range[1].GroupItems[i].Left;
                            float rl = range[1].GroupItems[i + 1].Left;
                            float Rt = range[1].GroupItems[i].Top;
                            float rt = range[1].GroupItems[i + 1].Top;
                            float a = rl + r - Rl - R;
                            float b = rt + r - Rt - R;
                            float c = R + r;
                            if (r + rl == R + Rl)
                            {
                                if (r + rt < R + Rt)
                                {
                                    range[1].GroupItems[i + 1].Top = Rt - 2 * r;
                                }

                                if (r + rt >= R + Rt)
                                {
                                    range[1].GroupItems[i + 1].Top = Rt + 2 * R;
                                }
                            }
                            else
                            {
                                if (2 * r + rl < Rl || 2 * R + Rl < rl)
                                {
                                    forms.MessageBox.Show("超出范围，不能垂直贴合");
                                }
                                else
                                {
                                    if (r + rt >= R + Rt)
                                    {
                                        float b1 = (float)Math.Sqrt(c * c - a * a);
                                        range[1].GroupItems[i + 1].Top = Rt + R + b1 - r;
                                    }
                                    if (r + rt < R + Rt)
                                    {
                                        float b1 = (float)Math.Sqrt(c * c - a * a);
                                        range[1].GroupItems[i + 1].Top = Rt + R - b1 - r;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("请至少选择两个形状");
                    }
                }
                else
                {
                    for (int i = 1; i < count; i++)
                    {
                        float R = range[i].Width / 2;
                        float r = range[i + 1].Width / 2;
                        float Rl = range[i].Left;
                        float rl = range[i + 1].Left;
                        float Rt = range[i].Top;
                        float rt = range[i + 1].Top;
                        float a = rl + r - Rl - R;
                        float b = rt + r - Rt - R;
                        float c = R + r;
                        if (r + rl == R + Rl)
                        {
                            if (r + rt < R + Rt)
                            {
                                range[i + 1].Top = Rt - 2 * r;
                            }

                            if (r + rt >= R + Rt)
                            {
                                range[i + 1].Top = Rt + 2 * R;
                            }
                        }
                        else
                        {
                            if (2 * r + rl < Rl || 2 * R + Rl < rl)
                            {
                                forms.MessageBox.Show("超出范围，不能垂直贴合");
                            }
                            else
                            {
                                if (r + rt >= R + Rt)
                                {
                                    float b1 = (float)Math.Sqrt(c * c - a * a);
                                    range[i + 1].Top = Rt + R + b1 - r;
                                }
                                if (r + rt < R + Rt)
                                {
                                    float b1 = (float)Math.Sqrt(c * c - a * a);
                                    range[i + 1].Top = Rt + R - b1 - r;
                                }
                            }
                        }
                    }
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                forms.MessageBox.Show("请至少选择三个形状");
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
                int count = range.Count;
                if (count == 1)
                {
                    if (range[1].Type == Office.MsoShapeType.msoGroup && range[1].GroupItems.Count > 2)
                    {
                        float minl = range[1].GroupItems[1].Left;
                        float maxl = range[1].GroupItems[1].Left;
                        for (int i = 2; i <= range[1].GroupItems.Count; i++)
                        {
                            float ol = range[1].GroupItems[i].Left;
                            minl = Math.Min(ol, minl);
                            maxl = Math.Max(ol, maxl);
                        }
                        float n1 = (maxl - minl) / (range[1].GroupItems.Count - 1);
                        for (int j = 1; j <= range[1].GroupItems.Count; j++)
                        {
                            range[1].GroupItems[j].Left = minl + n1 * (j - 1);
                        }
                    }
                    else if (range[1].Type == Office.MsoShapeType.msoGroup && range[1].GroupItems.Count == 2)
                    {
                        MessageBox.Show("组合内至少要有三个形状");
                    }
                    else
                    {
                        MessageBox.Show("请至少选中三个形状，或一个组合内至少要有三个形状");
                    }
                }
                else if (count == 2)
                {
                    MessageBox.Show("请选择至少三个形状");
                }
                else
                {
                    float minl = range[1].Left;
                    float maxl = range[1].Left;
                    for (int i = 2; i <= count; i++)
                    {
                        float ol = range[i].Left;
                        minl = Math.Min(ol, minl);
                        maxl = Math.Max(ol, maxl);
                    }
                    float n1 = (maxl - minl) / (count - 1);
                    for (int j = 1; j <= count; j++)
                    {
                        range[j].Left = minl + n1 * (j - 1);
                    }
                }   
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                forms.MessageBox.Show("请至少选择三个形状");
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
                int count = range.Count;
                if (count == 1)
                {
                    if (range[1].Type == Office.MsoShapeType.msoGroup && range[1].GroupItems.Count > 2)
                    {
                        float mint = range[1].GroupItems[1].Top;
                        float maxt = range[1].GroupItems[1].Top;
                        for (int i = 2; i <= range[1].GroupItems.Count; i++)
                        {
                            float ot = range[1].GroupItems[i].Top;
                            mint = Math.Min(ot, mint);
                            maxt = Math.Max(ot, maxt);
                        }
                        float n1 = (maxt - mint) / (range[1].GroupItems.Count - 1);
                        for (int j = 1; j <= range[1].GroupItems.Count; j++)
                        {
                            range[1].GroupItems[j].Top = mint + n1 * (j - 1);
                        }
                    }
                    else if (range[1].Type == Office.MsoShapeType.msoGroup && range[1].GroupItems.Count == 2)
                    {
                        MessageBox.Show("组合内至少要有三个形状");
                    }
                    else
                    {
                        MessageBox.Show("请至少选中三个形状，或一个组合内至少要有三个形状");
                    }
                }
                else if (count == 2)
                {
                    MessageBox.Show("请选择至少三个形状");
                }
                else
                {
                    float mint = range[1].Top;
                    float maxt = range[1].Top;
                    for (int i = 2; i <= count; i++)
                    {
                        float ot = range[i].Top;
                        mint = Math.Min(ot, mint);
                        maxt = Math.Max(ot, maxt);
                    }
                    float n1 = (maxt - mint) / (count - 1);
                    for (int j = 1; j <= count; j++)
                    {
                        range[j].Top = mint + n1 * (j - 1);
                    }
                }  
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                forms.MessageBox.Show("请至少选择三个形状");
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
                int count = range.Count;
                if (count == 1)
                {
                    if (range[1].Type == Office.MsoShapeType.msoGroup && range[1].GroupItems.Count > 2)
                    {
                        float nw = range[1].GroupItems[range[1].GroupItems.Count].Left - range[1].GroupItems[1].Left - range[1].GroupItems[1].Width;
                        float nw2 = range[1].GroupItems[range[1].GroupItems.Count].Left - range[1].GroupItems[1].Left - range[1].GroupItems[1].Width;
                        float cw = 0;
                        if (range[1].GroupItems.Count == 3)
                        {
                            if (range[1].GroupItems[2].Width >= nw)
                            {
                                range[1].GroupItems[2].Left = range[1].GroupItems[1].Left + range[1].GroupItems[1].Width + nw / 3;
                            }
                            else
                            {
                                nw = nw - range[1].GroupItems[2].Width;
                                range[1].GroupItems[2].Left = range[1].GroupItems[1].Left + range[1].GroupItems[1].Width + nw / 3;
                            }
                        }
                        else
                        {
                            List<float> widths = new List<float>();
                            foreach (PowerPoint.Shape item in range.GroupItems)
                            {
                                widths.Add(item.Width);
                            }

                            for (int j = 2; j < range[1].GroupItems.Count; j++)
                            {
                                PowerPoint.Shape item = range[1].GroupItems[j];
                                cw = cw + item.Width;
                                nw = nw - item.Width;
                            }

                            for (int k = 3; k < range[1].GroupItems.Count; k++)
                            {
                                float avgnw = 0;
                                if (cw > nw2)
                                {
                                    if (range[1].GroupItems.Count % 2 != 0)
                                    {
                                        avgnw = 2 * nw2 / (range[1].GroupItems.Count * range[1].GroupItems.Count);
                                    }
                                    else
                                    {
                                        avgnw = 2 * nw2 / (range[1].GroupItems.Count * (range[1].GroupItems.Count - 1));
                                    }
                                    range[1].GroupItems[2].Left = range[1].GroupItems[1].Left + range[1].GroupItems[1].Width + avgnw;
                                    range[1].GroupItems[k].Left = range[1].GroupItems[k - 1].Left + avgnw * (k - 1);
                                }
                                else
                                {
                                    if (range[1].GroupItems.Count % 2 != 0)
                                    {
                                        avgnw = 2 * nw / (range[1].GroupItems.Count * range[1].GroupItems.Count);

                                    }
                                    else
                                    {
                                        avgnw = 2 * nw / (range[1].GroupItems.Count * (range[1].GroupItems.Count - 1));
                                    }
                                    range[1].GroupItems[2].Left = range[1].GroupItems[1].Left + range[1].GroupItems[1].Width + avgnw;
                                    range[1].GroupItems[k].Left = range[1].GroupItems[k - 1].Left + range[1].GroupItems[k - 1].Width + avgnw * (k - 1);
                                }
                            } 
                        }
                    }
                    else if (range[1].Type == Office.MsoShapeType.msoGroup && range[1].GroupItems.Count == 2)
                    {
                        MessageBox.Show("组合内至少要有三个形状");
                    }
                    else
                    {
                        MessageBox.Show("请至少选中三个形状，或一个组合内至少要有三个形状");
                    }
                }
                else if (count == 2)
                {
                    MessageBox.Show("请选择至少三个形状");
                }
                else
                {
                    float nw = range[count].Left - range[1].Left - range[1].Width;
                    float nw2 = range[count].Left - range[1].Left - range[1].Width;
                    float cw = 0;
                    if (count == 3)
                    {
                        if (range[2].Width >= nw)
                        {
                            range[2].Left = range[1].Left + range[1].Width + nw / 3;
                        }
                        else
                        {
                            nw = nw - range[2].Width;
                            range[2].Left = range[1].Left + range[1].Width + nw / 3;
                        }
                    }
                    else
                    {
                        List<float> widths = new List<float>();
                        foreach (PowerPoint.Shape item in range)
                        {
                            widths.Add(item.Width);
                        }

                        for (int j = 2; j < count; j++)
                        {
                            PowerPoint.Shape item = range[j];
                            cw = cw + item.Width;
                            nw = nw - item.Width;
                        }

                        for (int k = 3; k < count; k++)
                        {
                            float avgnw = 0;
                            if (cw > nw2)
                            {
                                if (count % 2 != 0)
                                {
                                    avgnw = 2 * nw2 / (count * count);
                                }
                                else
                                {
                                    avgnw = 2 * nw2 / (count * (count - 1));
                                }
                                range[2].Left = range[1].Left + range[1].Width + avgnw;
                                range[k].Left = range[k - 1].Left + avgnw * (k - 1);
                            }
                            else
                            {
                                if (count % 2 != 0)
                                {
                                    avgnw = 2 * nw / (count * count);

                                }
                                else
                                {
                                    avgnw = 2 * nw / (count * (count - 1));
                                }
                                range[2].Left = range[1].Left + range[1].Width + avgnw;
                                range[k].Left = range[k - 1].Left + range[k - 1].Width + avgnw * (k - 1);
                            }
                        }
                    }
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                forms.MessageBox.Show("请至少选择三个形状");
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
                int count = range.Count;
                if (count == 1)
                {
                    if (range[1].Type == Office.MsoShapeType.msoGroup && range[1].GroupItems.Count > 2)
                    {
                        float nt = range[1].GroupItems[count].Top - range[1].GroupItems[1].Top - range[1].GroupItems[1].Height;
                        float nt2 = range[1].GroupItems[range[1].GroupItems.Count].Top - range[1].GroupItems[1].Top - range[1].GroupItems[1].Height;
                        float ct = 0;
                        if (range[1].GroupItems.Count == 3)
                        {
                            if (range[1].GroupItems[2].Height >= nt)
                            {
                                range[1].GroupItems[2].Top = range[1].GroupItems[1].Top + range[1].GroupItems[1].Height + nt / 3;
                            }
                            else
                            {
                                nt = nt - range[1].GroupItems[2].Height;
                                range[1].GroupItems[2].Top = range[1].GroupItems[1].Top + range[1].GroupItems[1].Height + nt / 3;
                            }
                        }
                        else
                        {
                            List<float> heights = new List<float>();
                            foreach (PowerPoint.Shape item in range.GroupItems)
                            {
                                heights.Add(item.Height);
                            }

                            for (int j = 2; j < range[1].GroupItems.Count; j++)
                            {
                                PowerPoint.Shape item = range[1].GroupItems[j];
                                ct = ct + item.Height;
                                nt = nt - item.Height;
                            }

                            for (int k = 3; k < range[1].GroupItems.Count; k++)
                            {
                                float avgnt = 0;
                                if (ct > nt2)
                                {
                                    if (range[1].GroupItems.Count % 2 != 0)
                                    {
                                        avgnt = 2 * nt2 / (range[1].GroupItems.Count * range[1].GroupItems.Count);
                                    }
                                    else
                                    {
                                        avgnt = 2 * nt2 / (range[1].GroupItems.Count * (range[1].GroupItems.Count - 1));
                                    }
                                    range[1].GroupItems[2].Top = range[1].GroupItems[1].Top + range[1].GroupItems[1].Height + avgnt;
                                    range[1].GroupItems[k].Top = range[1].GroupItems[k - 1].Top + avgnt * (k - 1);
                                }
                                else
                                {
                                    if (range[1].GroupItems.Count % 2 != 0)
                                    {
                                        avgnt = 2 * nt / (range[1].GroupItems.Count * range[1].GroupItems.Count);

                                    }
                                    else
                                    {
                                        avgnt = 2 * nt / (range[1].GroupItems.Count * (range[1].GroupItems.Count - 1));
                                    }
                                    range[1].GroupItems[2].Top = range[1].GroupItems[1].Top + range[1].GroupItems[1].Height + avgnt;
                                    range[1].GroupItems[k].Top = range[1].GroupItems[k - 1].Top + range[1].GroupItems[k - 1].Height + avgnt * (k - 1);
                                }
                            }
                        }
                    }
                    else if (range[1].Type == Office.MsoShapeType.msoGroup && range[1].GroupItems.Count == 2)
                    {
                        MessageBox.Show("组合内至少要有三个形状");
                    }
                }
                else if (count == 2)
                {
                    MessageBox.Show("请选择至少三个形状");
                }
                else
                {
                    float nt = range[count].Top - range[1].Top - range[1].Height;
                    float nt2 = range[count].Top - range[1].Top - range[1].Height;
                    float ct = 0;
                    if (count == 3)
                    {
                        if (range[2].Height >= nt)
                        {
                            range[2].Top = range[1].Top + range[1].Height + nt / 3;
                        }
                        else
                        {
                            nt = nt - range[2].Height;
                            range[2].Top = range[1].Top + range[1].Height + nt / 3;
                        }
                    }
                    else
                    {
                        List<float> heights = new List<float>();
                        foreach (PowerPoint.Shape item in range)
                        {
                            heights.Add(item.Height);
                        }

                        for (int j = 2; j < count; j++)
                        {
                            PowerPoint.Shape item = range[j];
                            ct = ct + item.Height;
                            nt = nt - item.Height;
                        }

                        for (int k = 3; k < count; k++)
                        {
                            float avgnt = 0;
                            if (ct > nt2)
                            {
                                if (count % 2 != 0)
                                {
                                    avgnt = 2 * nt2 / (count * count);
                                }
                                else
                                {
                                    avgnt = 2 * nt2 / (count * (count - 1));
                                }
                                range[2].Top = range[1].Top + range[1].Height + avgnt;
                                range[k].Top = range[k - 1].Top + avgnt * (k - 1);
                            }
                            else
                            {
                                if (count % 2 != 0)
                                {
                                    avgnt = 2 * nt / (count * count);

                                }
                                else
                                {
                                    avgnt = 2 * nt / (count * (count - 1));
                                }
                                range[2].Top = range[1].Top + range[1].Height + avgnt;
                                range[k].Top = range[k - 1].Top + range[k - 1].Height + avgnt * (k - 1);
                            }
                        }
                    }
                }  
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                forms.MessageBox.Show("请至少选择三个形状");
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
                int count = range.Count;
                if (count == 1)
                {
                    if (range[1].Type == Office.MsoShapeType.msoGroup && range[1].GroupItems.Count > 2)
                    {
                        float firmid = range[1].GroupItems[1].Width / 2 + range[1].GroupItems[1].Left;
                        float endmid = range[1].GroupItems[range[1].GroupItems.Count].Width / 2 + range[1].GroupItems[range[1].GroupItems.Count].Left;
                        float nmid1 = (endmid - firmid) / (range[1].GroupItems.Count - 1);
                        for (int j = 2; j < range[1].GroupItems.Count; j++)
                        {
                            range[1].GroupItems[j].Left = range[1].GroupItems[j - 1].Width / 2 + range[1].GroupItems[j - 1].Left + nmid1 - range[1].GroupItems[j].Width / 2;
                        }
                    }
                    else if (range[1].Type == Office.MsoShapeType.msoGroup && range[1].GroupItems.Count == 2)
                    {
                        MessageBox.Show("组合内至少要有三个形状");
                    }
                    else
                    {
                        MessageBox.Show("请至少选中三个形状，或一个组合内至少要有三个形状");
                    }
                }
                else if (count == 2)
                {
                    MessageBox.Show("请选择至少三个形状");
                }
                else
                {
                    float firmid = range[1].Width / 2 + range[1].Left;
                    float endmid = range[count].Width / 2 + range[count].Left;
                    float nmid1 = (endmid - firmid) / (count - 1);
                    for (int j = 2; j < count; j++)
                    {
                        range[j].Left = range[j - 1].Width / 2 + range[j - 1].Left + nmid1 - range[j].Width / 2;
                    }
                }
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                forms.MessageBox.Show("请至少选择三个形状");
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
                int count = range.Count;
                if (count == 1)
                {
                    if (range[1].Type == Office.MsoShapeType.msoGroup && range[1].GroupItems.Count > 2)
                    {
                        float firmid = range[1].GroupItems[1].Height / 2 + range[1].GroupItems[1].Top;
                        float endmid = range[1].GroupItems[range[1].GroupItems.Count].Height / 2 + range[1].GroupItems[range[1].GroupItems.Count].Top;
                        float nmid1 = (endmid - firmid) / (range[1].GroupItems.Count - 1);
                        for (int j = 2; j < range[1].GroupItems.Count; j++)
                        {
                            range[1].GroupItems[j].Top = range[1].GroupItems[j - 1].Height / 2 + range[1].GroupItems[j - 1].Top + nmid1 - range[1].GroupItems[j].Height / 2;
                        }
                    }
                    else if (range[1].Type == Office.MsoShapeType.msoGroup && range[1].GroupItems.Count == 2)
                    {
                        MessageBox.Show("组合内至少要有三个形状");
                    }
                }
                else if (count == 2)
                {
                    MessageBox.Show("请选中至少三个形状");
                }
                else
                {
                    float firmid = range[1].Height / 2 + range[1].Top;
                    float endmid = range[count].Height / 2 + range[count].Top;
                    float nmid1 = (endmid - firmid) / (count - 1);
                    for (int j = 2; j < count; j++)
                    {
                        range[j].Top = range[j - 1].Height / 2 + range[j - 1].Top + nmid1 - range[j].Height / 2;
                    }
                }
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            Align_More.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button20.Enabled = true;
        }

    }
}
