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
    public partial class Table_Calculation : Form
    {
        public Table_Calculation()
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
            Table_Calculation.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button254.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Shape shape = sel.ShapeRange[1];
                if (shape.Type == Office.MsoShapeType.msoTable)
                {
                    try
                    {
                        double YinZi = double.Parse(textBox1.Text.Trim());
                        List<int> numi = new List<int>();
                        List<int> numj = new List<int>();

                        for (int i = 1; i <= shape.Table.Columns.Count; i++)
                        {
                            for (int j = 1; j <= shape.Table.Rows.Count; j++)
                            {
                                if (shape.Table.Columns[i].Cells[j].Selected)
                                {
                                    numi.Add(i);
                                    numj.Add(j);
                                }
                            }
                        }

                        numi = numi.Distinct().ToList();
                        numj = numj.Distinct().ToList();
                        int clcnt = numi.Count();
                        int rwcnt = numj.Count();
                        double value = 0;

                        for (int i = 0; i < clcnt; i++)
                        {
                            for (int j = 0; j < rwcnt; j++)
                            {
                                try
                                {
                                    if (radioButton1.Checked)
                                    {
                                        value = double.Parse(shape.Table.Columns[numi[i]].Cells[numj[j]].Shape.TextFrame.TextRange.Text.Trim()) + YinZi;
                                        shape.Table.Columns[numi[i]].Cells[numj[j]].Shape.TextFrame.TextRange.Text = value.ToString();
                                        value = 0;
                                    }
                                    else if (radioButton2.Checked)
                                    {
                                        value = double.Parse(shape.Table.Columns[numi[i]].Cells[numj[j]].Shape.TextFrame.TextRange.Text.Trim()) - YinZi;
                                        shape.Table.Columns[numi[i]].Cells[numj[j]].Shape.TextFrame.TextRange.Text = value.ToString();
                                        value = 0;
                                    }
                                    else if (radioButton3.Checked)
                                    {
                                        value = double.Parse(shape.Table.Columns[numi[i]].Cells[numj[j]].Shape.TextFrame.TextRange.Text.Trim()) * YinZi;
                                        shape.Table.Columns[numi[i]].Cells[numj[j]].Shape.TextFrame.TextRange.Text = value.ToString();
                                        value = 0;
                                    }
                                    else if (radioButton4.Checked)
                                    {
                                        value = double.Parse(shape.Table.Columns[numi[i]].Cells[numj[j]].Shape.TextFrame.TextRange.Text.Trim()) / YinZi;
                                        shape.Table.Columns[numi[i]].Cells[numj[j]].Shape.TextFrame.TextRange.Text = value.ToString();
                                        value = 0;
                                    }
                                    if (radioButton5.Checked)
                                    {
                                        value = double.Parse(shape.Table.Columns[numi[i]].Cells[numj[j]].Shape.TextFrame.TextRange.Text.Trim());
                                        value = Math.Pow(value, YinZi);
                                        shape.Table.Columns[numi[i]].Cells[numj[j]].Shape.TextFrame.TextRange.Text = value.ToString();
                                        value = 0;
                                    }
                                    else if (radioButton6.Checked)
                                    {
                                        value = double.Parse(shape.Table.Columns[numi[i]].Cells[numj[j]].Shape.TextFrame.TextRange.Text.Trim());
                                        value = Math.Pow(value, 1 / YinZi);
                                        shape.Table.Columns[numi[i]].Cells[numj[j]].Shape.TextFrame.TextRange.Text = value.ToString();
                                        value = 0;
                                    }
                                    else if (radioButton7.Checked)
                                    {
                                        value = double.Parse(shape.Table.Columns[numi[i]].Cells[numj[j]].Shape.TextFrame.TextRange.Text.Trim());
                                        value = Math.Round(value,(int)YinZi, MidpointRounding.AwayFromZero);
                                        shape.Table.Columns[numi[i]].Cells[numj[j]].Shape.TextFrame.TextRange.Text = value.ToString();
                                        value = 0;
                                    }
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }
                    }
                    catch
                    {
                        MessageBox.Show("因子的数值有误，请核对");
                        throw;
                    }
                }
                else
                {
                    MessageBox.Show("请选中表格、单元格区域");
                }
            }
            else
            {
                MessageBox.Show("请选中表格、单元格区域");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Shape shape = sel.ShapeRange[1];
                if (shape.Type == Office.MsoShapeType.msoTable && shape.Table.Rows.Count > 1)
                {
                    try
                    {
                        float step = float.Parse(textBox2.Text.Trim());
                        int begin = 0;
                        float numbegin = 0;
                        float value = 0;
                        for (int i = 1; i <= shape.Table.Columns.Count; i++)
                        {
                            int n = 0;
                            for (int j = 1; j <= shape.Table.Rows.Count; j++)
                            {
                                if (shape.Table.Columns[i].Cells[j].Selected)
                                {
                                    if (begin == 0)
                                    {
                                        try
                                        {
                                            numbegin = float.Parse(shape.Table.Columns[i].Cells[j].Shape.TextFrame.TextRange.Text.Trim());
                                            begin = 1;
                                            n = 1;
                                        }
                                        catch
                                        {
                                            numbegin = 1;
                                            begin = 1;
                                            continue;
                                        }
                                    }
                                    else
                                    {
                                        try
                                        {
                                            value = numbegin + step * n;
                                            shape.Table.Columns[i].Cells[j].Shape.TextFrame.TextRange.Text = value.ToString();
                                            n += 1;
                                        }
                                        catch
                                        {
                                            continue;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        MessageBox.Show("步长格式有误，请输入纯数字");
                    } 
                }
                else
                {
                    MessageBox.Show("请先选中 指定列 的单元格，且第一个单元格已经输入初始数值");
                }
            }
            else
            {
                MessageBox.Show("请先选中 指定列 的单元格，且第一个单元格已经输入初始数值");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Shape shape = sel.ShapeRange[1];
                if (shape.Type == Office.MsoShapeType.msoTable && shape.Table.Columns.Count > 1)
                {
                    try
                    {
                        float step = float.Parse(textBox2.Text.Trim());
                        int begin = 0;
                        float numbegin = 0;
                        float value = 0;
                        for (int i = 1; i <= shape.Table.Rows.Count; i++)
                        {
                            int n = 0;
                            for (int j = 1; j <= shape.Table.Columns.Count; j++)
                            {
                                if (shape.Table.Rows[i].Cells[j].Selected)
                                {
                                    if (begin == 0)
                                    {
                                        try
                                        {
                                            numbegin = float.Parse(shape.Table.Rows[i].Cells[j].Shape.TextFrame.TextRange.Text.Trim());
                                            begin = 1;
                                            n = 1;
                                        }
                                        catch
                                        {
                                            numbegin = 1;
                                            begin = 1;
                                            continue;
                                        }
                                    }
                                    else
                                    {
                                        try
                                        {
                                            value = numbegin + step * n;
                                            shape.Table.Rows[i].Cells[j].Shape.TextFrame.TextRange.Text = value.ToString();
                                            n += 1;
                                        }
                                        catch
                                        {
                                            continue;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        MessageBox.Show("步长格式有误，请输入纯数字");
                    }
                }
                else
                {
                    MessageBox.Show("请先选中 指定行 的单元格，且第一个单元格已经输入初始数值");
                }
            }
            else
            {
                MessageBox.Show("请先选中 指定行 的单元格，且第一个单元格已经输入初始数值");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Shape shape = sel.ShapeRange[1];
                if (shape.Type == Office.MsoShapeType.msoTable)
                {
                    try
                    {
                        float begin = float.Parse(textBox3.Text.Trim());
                        float end = float.Parse(textBox4.Text.Trim());
                        Random rand = new Random();
                        int ran1 = 0;
                        double ran2 = 0;
                        for (int i = 1; i <= shape.Table.Rows.Count; i++)
                        {
                            for (int j = 1; j <= shape.Table.Columns.Count; j++)
                            {
                                if (shape.Table.Rows[i].Cells[j].Selected)
                                {
                                    if (begin == (int)begin && end == (int)end)
                                    {
                                        ran1 = rand.Next((int)begin * 1000000, (int)end * 1000000) / 1000000;
                                        shape.Table.Rows[i].Cells[j].Shape.TextFrame.TextRange.Text = ran1.ToString();
                                    }
                                    else
                                    {
                                        ran2 = (double)rand.Next((int)begin * 1000000, (int)end * 1000000) / 1000000;
                                        shape.Table.Rows[i].Cells[j].Shape.TextFrame.TextRange.Text = ran2.ToString();

                                    }  
                                }
                            }
                        }
                    }
                    catch
                    {
                        MessageBox.Show("随机范围值的格式有误，请输入纯数字");
                    }
                }
                else
                {
                    MessageBox.Show("请先选中 指定行 的单元格，且第一个单元格已经输入初始数值");
                }
            }
        }

    }
}
