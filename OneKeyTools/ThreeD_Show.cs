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
using System.Runtime.InteropServices;

namespace OneKeyTools
{
    public partial class ThreeD_Show : Form
    {
        public ThreeD_Show()
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
            if (radioButton1.Checked)
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.Shapes shapes = slide.Shapes;
                foreach (PowerPoint.Shape item in shapes)
                {
                    if (item.Name == comboBox1.SelectedItem.ToString())
                    {
                        float n = float.Parse(textBox1.Text.Trim());
                        if (item.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                        {
                            foreach (PowerPoint.Shape item2 in item.GroupItems)
                            {
                                item2.ThreeD.IncrementRotationX(n);
                            }
                        }
                        else
                        {
                            item.ThreeD.IncrementRotationX(n);
                        }
                    }
                }
            }
            if (radioButton2.Checked)
            {
                PowerPoint.Shapes shapes = app.SlideShowWindows[1].View.Slide.Shapes;
                foreach (PowerPoint.Shape item in shapes)
                {
                    if (item.Name == comboBox1.SelectedItem.ToString())
                    {
                        float n = float.Parse(textBox1.Text.Trim());
                        if (item.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                        {
                            foreach (PowerPoint.Shape item2 in item.GroupItems)
                            {
                                item2.ThreeD.IncrementRotationX(n);
                            }
                        }
                        else
                        {
                            item.ThreeD.IncrementRotationX(n);
                        }
                    }
                }
            }  
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ThreeD_Show.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button142.Enabled = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            if (radioButton1.Checked)
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.Shapes shapes = slide.Shapes;
                foreach (PowerPoint.Shape item in shapes)
                {
                    if (item.ThreeD.Visible == Microsoft.Office.Core.MsoTriState.msoTrue || (item.Type == Microsoft.Office.Core.MsoShapeType.msoGroup && (item.GroupItems[1].ThreeD.Visible == Microsoft.Office.Core.MsoTriState.msoTrue || item.GroupItems[2].ThreeD.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)))
                    {
                        comboBox1.Items.Add(item.Name);
                    }
                }
            }
            if (radioButton2.Checked)
            {
                PowerPoint.Shapes shapes = app.SlideShowWindows[1].View.Slide.Shapes;
                foreach (PowerPoint.Shape item in shapes)
                {
                    if (item.ThreeD.Visible == Microsoft.Office.Core.MsoTriState.msoTrue || (item.Type == Microsoft.Office.Core.MsoShapeType.msoGroup && (item.GroupItems[1].ThreeD.Visible == Microsoft.Office.Core.MsoTriState.msoTrue || item.GroupItems[2].ThreeD.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)))
                    {
                        comboBox1.Items.Add(item.Name);
                    }
                }
            }
            if (comboBox1.Items.Count == 0)
            {
                MessageBox.Show("没有在当前页面中找到三维形状，请确认后再刷新");
                comboBox1.Enabled = false;
                textBox1.Enabled = false;
                button1.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;
                checkBox1.Enabled = false;
                checkBox2.Enabled = false;
                checkBox3.Enabled = false;
                checkBox4.Enabled = false;
                textBox2.Enabled = false;
                button6.Enabled = false;
                label3.Enabled = false;
                comboBox1.Text = "";
            }
            else
            {
                comboBox1.Enabled = true;
                comboBox1.Text = "";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.Shapes shapes = slide.Shapes;
                foreach (PowerPoint.Shape item in shapes)
                {
                    if (item.Name == comboBox1.SelectedItem.ToString())
                    {
                        float n = float.Parse(textBox1.Text.Trim());
                        if (item.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                        {
                            foreach (PowerPoint.Shape item2 in item.GroupItems)
                            {
                                item2.ThreeD.IncrementRotationY(n);
                            }
                        }
                        else
                        {
                            item.ThreeD.IncrementRotationY(n);
                        }
                    }
                }
            }
            if (radioButton2.Checked)
            {
                PowerPoint.Shapes shapes = app.SlideShowWindows[1].View.Slide.Shapes;
                foreach (PowerPoint.Shape item in shapes)
                {
                    if (item.Name == comboBox1.SelectedItem.ToString())
                    {
                        float n = float.Parse(textBox1.Text.Trim());
                        if (item.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                        {
                            foreach (PowerPoint.Shape item2 in item.GroupItems)
                            {
                                item2.ThreeD.IncrementRotationY(n);
                            }
                        }
                        else
                        {
                            item.ThreeD.IncrementRotationY(n);
                        }
                    }
                }
            }  
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.Shapes shapes = slide.Shapes;
                foreach (PowerPoint.Shape item in shapes)
                {
                    if (item.Name == comboBox1.SelectedItem.ToString())
                    {
                        float n = float.Parse(textBox1.Text.Trim());
                        if (item.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                        {
                            foreach (PowerPoint.Shape item2 in item.GroupItems)
                            {
                                item2.ThreeD.IncrementRotationZ(n);
                            }
                        }
                        else
                        {
                            item.ThreeD.IncrementRotationZ(n);
                        }
                    }
                }
            }
            if (radioButton2.Checked)
            {
                PowerPoint.Shapes shapes = app.SlideShowWindows[1].View.Slide.Shapes;
                foreach (PowerPoint.Shape item in shapes)
                {
                    if (item.Name == comboBox1.SelectedItem.ToString())
                    {
                        float n = float.Parse(textBox1.Text.Trim());
                        if (item.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                        {
                            foreach (PowerPoint.Shape item2 in item.GroupItems)
                            {
                                item2.ThreeD.IncrementRotationZ(n);
                            }
                        }
                        else
                        {
                            item.ThreeD.IncrementRotationZ(n);
                        }
                    }
                }
            }  
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                button1.Enabled = false;
                button2.Enabled = false;
                button3.Enabled = false;
                checkBox2.Enabled = true;
                checkBox3.Enabled = true;
                checkBox4.Enabled = true;
                button6.Enabled = true;
                label3.Enabled = true;
                textBox2.Enabled = true;
            }
            else
            {
                button1.Enabled = true;
                button2.Enabled = true;
                button3.Enabled = true;
                checkBox2.Enabled = false;
                checkBox3.Enabled = false;
                checkBox4.Enabled = false;
                button6.Enabled = false;
                timer1.Enabled = false;
                label3.Enabled = false;
                textBox2.Enabled = false;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (timer1.Enabled == false)
            {
                timer1.Enabled = true;
            }
            else
            {
                timer1.Enabled = false;
            } 
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (!checkBox2.Checked && !checkBox3.Checked && !checkBox4.Checked)
            {
                timer1.Enabled = false;
                MessageBox.Show("请先勾选“X”、“Y”、“Z”中的一个或多个");
            }
            else
            {
                if (checkBox2.Checked)
                {
                    if (radioButton1.Checked)
                    {
                        PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                        PowerPoint.Shapes shapes = slide.Shapes;
                        foreach (PowerPoint.Shape item in shapes)
                        {
                            if (item.Name == comboBox1.SelectedItem.ToString())
                            {
                                float n = float.Parse(textBox1.Text.Trim());
                                if (item.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                                {
                                    foreach (PowerPoint.Shape item2 in item.GroupItems)
                                    {
                                        item2.ThreeD.IncrementRotationX(n);
                                    }
                                }
                                else
                                {
                                    item.ThreeD.IncrementRotationX(n);
                                }
                            }
                        }
                    }
                    if (radioButton2.Checked)
                    {
                        PowerPoint.Shapes shapes = app.SlideShowWindows[1].View.Slide.Shapes;
                        foreach (PowerPoint.Shape item in shapes)
                        {
                            if (item.Name == comboBox1.SelectedItem.ToString())
                            {
                                float n = float.Parse(textBox1.Text.Trim());
                                if (item.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                                {
                                    foreach (PowerPoint.Shape item2 in item.GroupItems)
                                    {
                                        item2.ThreeD.IncrementRotationX(n);
                                    }
                                }
                                else
                                {
                                    item.ThreeD.IncrementRotationX(n);
                                }
                            }
                        }
                    }
                }
                if (checkBox3.Checked)
                {
                    if (radioButton1.Checked)
                    {
                        PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                        PowerPoint.Shapes shapes = slide.Shapes;
                        foreach (PowerPoint.Shape item in shapes)
                        {
                            if (item.Name == comboBox1.SelectedItem.ToString())
                            {
                                float n = float.Parse(textBox1.Text.Trim());
                                if (item.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                                {
                                    foreach (PowerPoint.Shape item2 in item.GroupItems)
                                    {
                                        item2.ThreeD.IncrementRotationY(n);
                                    }
                                }
                                else
                                {
                                    item.ThreeD.IncrementRotationY(n);
                                }
                            }
                        }
                    }
                    if (radioButton2.Checked)
                    {
                        PowerPoint.Shapes shapes = app.SlideShowWindows[1].View.Slide.Shapes;
                        foreach (PowerPoint.Shape item in shapes)
                        {
                            if (item.Name == comboBox1.SelectedItem.ToString())
                            {
                                float n = float.Parse(textBox1.Text.Trim());
                                if (item.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                                {
                                    foreach (PowerPoint.Shape item2 in item.GroupItems)
                                    {
                                        item2.ThreeD.IncrementRotationY(n);
                                    }
                                }
                                else
                                {
                                    item.ThreeD.IncrementRotationY(n);
                                }
                            }
                        }
                    }  
                }
                if (checkBox4.Checked)
                {
                    if (radioButton1.Checked)
                    {
                        PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                        PowerPoint.Shapes shapes = slide.Shapes;
                        foreach (PowerPoint.Shape item in shapes)
                        {
                            if (item.Name == comboBox1.SelectedItem.ToString())
                            {
                                float n = float.Parse(textBox1.Text.Trim());
                                if (item.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                                {
                                    foreach (PowerPoint.Shape item2 in item.GroupItems)
                                    {
                                        item2.ThreeD.IncrementRotationZ(n);
                                    }
                                }
                                else
                                {
                                    item.ThreeD.IncrementRotationZ(n);
                                }
                            }
                        }
                    }
                    if (radioButton2.Checked)
                    {
                        PowerPoint.Shapes shapes = app.SlideShowWindows[1].View.Slide.Shapes;
                        foreach (PowerPoint.Shape item in shapes)
                        {
                            if (item.Name == comboBox1.SelectedItem.ToString())
                            {
                                float n = float.Parse(textBox1.Text.Trim());
                                if (item.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                                {
                                    foreach (PowerPoint.Shape item2 in item.GroupItems)
                                    {
                                        item2.ThreeD.IncrementRotationZ(n);
                                    }
                                }
                                else
                                {
                                    item.ThreeD.IncrementRotationZ(n);
                                }
                            }
                        }
                    }
                }
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            int n = int.Parse(textBox2.Text.Trim());
            timer1.Interval = n;
        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem as string != "")
            {
                textBox1.Enabled = true;
                button1.Enabled = true;
                button2.Enabled = true;
                button3.Enabled = true;
                checkBox1.Enabled = true;
            }
        }

        class HotKey //http://www.cnblogs.com/Asa-Zhu/archive/2012/11/08/2761104.html
        {
            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool RegisterHotKey(IntPtr hWnd, int id, KeyModifiers fsModifiers, Keys vk);
            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool UnregisterHotKey(IntPtr hWnd,int id);
            [Flags()]
            public enum KeyModifiers
            {
                None = 0,
                Alt = 1,
                Ctrl = 2,
                Shift = 4,
                WindowsKey = 8
            }
        }

        private void timedj_Activated(object sender, EventArgs e)
        {
            HotKey.RegisterHotKey(Handle, 100, HotKey.KeyModifiers.Ctrl | HotKey.KeyModifiers.Shift | HotKey.KeyModifiers.Alt, Keys.H);
        }

        private void timedj_Leave(object sender, EventArgs e)
        {
            HotKey.UnregisterHotKey(Handle, 100);
        }

        protected override void WndProc(ref Message m)
        {
            const int WM_HOTKEY = 0x0312;
            switch (m.Msg)
            {
                case WM_HOTKEY:
                    switch (m.WParam.ToInt32())
                    {
                        case 100:
                            if (this.WindowState == FormWindowState.Normal)
                            {
                                this.WindowState = FormWindowState.Minimized;
                            }
                            else
	                        {
                                this.WindowState = FormWindowState.Normal;
	                        }
                            break;
                    }
                    break;
            }
            base.WndProc(ref m);
        }

    }
}
