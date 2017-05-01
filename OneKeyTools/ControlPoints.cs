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
    public partial class ControlPoints : Form
    {
        public ControlPoints()
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

        private void timer1_Tick(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionNone)
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
                PowerPoint.Shape shape = range[1];
                PowerPoint.Adjustments adj = shape.Adjustments;
                if (range[1].Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                {
                    adj = shape.GroupItems[1].Adjustments;
                }
                int acount = adj.Count;
                if (acount == 1)
                {
                    label1.Visible = true;
                    label2.Visible = false;
                    label3.Visible = false;
                    label4.Visible = false;
                    label6.Visible = true;
                    label7.Visible = false;
                    label8.Visible = false;
                    label9.Visible = false;
                    textBox1.Visible = true;
                    textBox2.Visible = false;
                    textBox3.Visible = false;
                    textBox4.Visible = false;
                    textBox5.Visible = true;
                    textBox6.Visible = false;
                    textBox7.Visible = false;
                    textBox8.Visible = false;
                    checkBox1.Visible = true;
                    checkBox2.Visible = false;
                    checkBox3.Visible = false;
                    checkBox4.Visible = false;
                    label6.Text = adj[1].ToString();
                }
                if (acount == 2)
                {
                    label1.Visible = true;
                    label2.Visible = true;
                    label3.Visible = false;
                    label4.Visible = false;
                    label6.Visible = true;
                    label7.Visible = true;
                    label8.Visible = false;
                    label9.Visible = false;
                    textBox1.Visible = true;
                    textBox2.Visible = true;
                    textBox3.Visible = false;
                    textBox4.Visible = false;
                    textBox5.Visible = true;
                    textBox6.Visible = true;
                    textBox7.Visible = false;
                    textBox8.Visible = false;
                    checkBox1.Visible = true;
                    checkBox2.Visible = true;
                    checkBox3.Visible = false;
                    checkBox4.Visible = false;
                    label6.Text = adj[1].ToString();
                    label7.Text = adj[2].ToString();
                }
                if (acount == 3)
                {
                    label1.Visible = true;
                    label2.Visible = true;
                    label3.Visible = true;
                    label4.Visible = false;
                    label6.Visible = true;
                    label7.Visible = true;
                    label8.Visible = true;
                    label9.Visible = false;
                    textBox1.Visible = true;
                    textBox2.Visible = true;
                    textBox3.Visible = true;
                    textBox4.Visible = false;
                    textBox5.Visible = true;
                    textBox6.Visible = true;
                    textBox7.Visible = true;
                    textBox8.Visible = false;
                    checkBox1.Visible = true;
                    checkBox2.Visible = true;
                    checkBox3.Visible = true;
                    checkBox4.Visible = false;
                    label6.Text = adj[1].ToString();
                    label7.Text = adj[2].ToString();
                    label8.Text = adj[3].ToString();
                }
                if (acount == 4)
                {
                    label1.Visible = true; 
                    label2.Visible = true;
                    label3.Visible = true;
                    label4.Visible = true;
                    label6.Visible = true;
                    label7.Visible = true;
                    label8.Visible = true;
                    label9.Visible = true;
                    textBox1.Visible = true;
                    textBox2.Visible = true;
                    textBox3.Visible = true;
                    textBox4.Visible = true;
                    textBox5.Visible = true;
                    textBox6.Visible = true;
                    textBox7.Visible = true;
                    textBox8.Visible = true;
                    checkBox1.Visible = true;
                    checkBox2.Visible = true;
                    checkBox3.Visible = true;
                    checkBox4.Visible = true;
                    label6.Text = adj[1].ToString();
                    label7.Text = adj[2].ToString();
                    label8.Text = adj[3].ToString();
                    label9.Text = adj[4].ToString();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                MessageBox.Show("请先选中一个带控点的矢量形状");
            }
            else
            {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                int count = range.Count;
                if (!checkBox1.Checked && !checkBox2.Checked && !checkBox3.Checked && !checkBox4.Checked)
                {
                    for (int i = 1; i <= count; i++)
                    {
                        PowerPoint.Shape shape = range[i];
                        if (range[i].Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                        {
                            for (int j = 1; j <= range[i].GroupItems.Count; j++)
                            {
                                PowerPoint.Shape gshape = range[i].GroupItems[j];
                                PowerPoint.Adjustments adj = gshape.Adjustments;
                                int acount = adj.Count;
                                if (acount == 1)
                                {
                                    adj[1] = float.Parse(textBox1.Text);
                                }
                                if (acount == 2)
                                {
                                    adj[1] = float.Parse(textBox1.Text);
                                    adj[2] = float.Parse(textBox2.Text);
                                }
                                if (acount == 3)
                                {
                                    adj[1] = float.Parse(textBox1.Text);
                                    adj[2] = float.Parse(textBox2.Text);
                                    adj[3] = float.Parse(textBox3.Text);
                                }
                                if (acount == 4)
                                {
                                    adj[1] = float.Parse(textBox1.Text);
                                    adj[2] = float.Parse(textBox2.Text);
                                    adj[3] = float.Parse(textBox3.Text);
                                    adj[4] = float.Parse(textBox4.Text);
                                }
                            }
                        }
                        else
                        {
                            PowerPoint.Adjustments adj = shape.Adjustments;
                            int acount = adj.Count;
                            if (acount == 1)
                            {
                                adj[1] = float.Parse(textBox1.Text);
                            }
                            if (acount == 2)
                            {
                                adj[1] = float.Parse(textBox1.Text);
                                adj[2] = float.Parse(textBox2.Text);
                            }
                            if (acount == 3)
                            {
                                adj[1] = float.Parse(textBox1.Text);
                                adj[2] = float.Parse(textBox2.Text);
                                adj[3] = float.Parse(textBox3.Text);
                            }
                            if (acount == 4)
                            {
                                adj[1] = float.Parse(textBox1.Text);
                                adj[2] = float.Parse(textBox2.Text);
                                adj[3] = float.Parse(textBox3.Text);
                                adj[4] = float.Parse(textBox4.Text);
                            }
                        }
                    }
                }
                else
                {
                    if(count >= 3)
                    {
                        float n1 = 0; float n2 = 0; float n3 = 0; float n4 = 0;
                        if (checkBox1.Checked)
                        {
                            n1 = (float.Parse(textBox5.Text) - float.Parse(textBox1.Text)) / ((float)count - 1);
                        }
                        if (checkBox2.Checked)
                        {
                            n2 = (float.Parse(textBox6.Text) - float.Parse(textBox2.Text)) / ((float)count - 1);
                        }
                        if (checkBox3.Checked)
                        {
                            n3 = (float.Parse(textBox7.Text) - float.Parse(textBox3.Text)) / ((float)count - 1);
                        }
                        if (checkBox4.Checked)
                        {
                            n4 = (float.Parse(textBox8.Text) - float.Parse(textBox4.Text)) / ((float)count - 1);
                        }
                        for (int i = 1; i <= count; i++)
                        {
                            PowerPoint.Shape shape = range[i];
                            if (range[i].Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                            {
                                for (int j = 1; j <= range[i].GroupItems.Count; j++)
                                {
                                    PowerPoint.Shape gshape = range[i].GroupItems[j];
                                    PowerPoint.Adjustments adj = gshape.Adjustments;
                                    int acount = adj.Count;
                                    if (acount == 1)
                                    {
                                        adj[1] = float.Parse(textBox1.Text) + n1 * (i - 1);
                                    }
                                    if (acount == 2)
                                    {
                                        adj[1] = float.Parse(textBox1.Text) + n1 * (i - 1);
                                        adj[2] = float.Parse(textBox2.Text) + n2 * (i - 1);
                                    }
                                    if (acount == 3)
                                    {
                                        adj[1] = float.Parse(textBox1.Text) + n1 * (i - 1);
                                        adj[2] = float.Parse(textBox2.Text) + n2 * (i - 1);
                                        adj[3] = float.Parse(textBox3.Text) + n3 * (i - 1);
                                    }
                                    if (acount == 4)
                                    {
                                        adj[1] = float.Parse(textBox1.Text) + n1 * (i - 1);
                                        adj[2] = float.Parse(textBox2.Text) + n2 * (i - 1);
                                        adj[3] = float.Parse(textBox3.Text) + n3 * (i - 1);
                                        adj[4] = float.Parse(textBox4.Text) + n4 * (i - 1);
                                    }
                                }
                            }
                            else
                            {
                                PowerPoint.Adjustments adj = shape.Adjustments;
                                int acount = adj.Count;
                                if (acount == 1)
                                {
                                    adj[1] = float.Parse(textBox1.Text) + n1 * (i - 1); ;
                                }
                                if (acount == 2)
                                {
                                    adj[1] = float.Parse(textBox1.Text) + n1 * (i - 1); ;
                                    adj[2] = float.Parse(textBox2.Text) + n2 * (i - 1);
                                }
                                if (acount == 3)
                                {
                                    adj[1] = float.Parse(textBox1.Text) + n1 * (i - 1);
                                    adj[2] = float.Parse(textBox2.Text) + n2 * (i - 1);
                                    adj[3] = float.Parse(textBox3.Text) + n3 * (i - 1);
                                }
                                if (acount == 4)
                                {
                                    adj[1] = float.Parse(textBox1.Text) + n1 * (i - 1);
                                    adj[2] = float.Parse(textBox2.Text) + n2 * (i - 1);
                                    adj[3] = float.Parse(textBox3.Text) + n3 * (i - 1);
                                    adj[4] = float.Parse(textBox4.Text) + n4 * (i - 1);
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("控点递进需要至少3个带控点的形状");
                    }
                } 
            }
        }

        private void label6_TextChanged(object sender, EventArgs e)
        {
            textBox1.Text = label6.Text;
        }

        private void label7_TextChanged(object sender, EventArgs e)
        {
            textBox2.Text = label7.Text;
        }

        private void label8_TextChanged(object sender, EventArgs e)
        {
            textBox3.Text = label8.Text;
        }

        private void label9_TextChanged(object sender, EventArgs e)
        {
            textBox4.Text = label9.Text;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ControlPoints.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button30.Enabled = true;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                textBox5.Enabled = true;
            }
            else
            {
                textBox5.Enabled = false;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                textBox6.Enabled = true;
            }
            else
            {
                textBox6.Enabled = false;
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                textBox7.Enabled = true;
            }
            else
            {
                textBox7.Enabled = false;
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                textBox8.Enabled = true;
            }
            else
            {
                textBox8.Enabled = false;
            }
        }
    }
}
