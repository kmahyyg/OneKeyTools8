using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OneKeyTools
{
    public partial class OK_Settings : Form
    {
        public OK_Settings()
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

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                Globals.Ribbons.Ribbon1.group2.Visible = true;
                Properties.Settings.Default.g1 = 1;
                Properties.Settings.Default.Save();
            }
            else
            {
                Globals.Ribbons.Ribbon1.group2.Visible = false;
                Properties.Settings.Default.g1 = 0;
                Properties.Settings.Default.Save();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OK_Settings.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button127.Enabled = true;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                Globals.Ribbons.Ribbon1.group3.Visible = true;
                Properties.Settings.Default.g2 = 1;
                Properties.Settings.Default.Save();
            }
            else
            {
                Globals.Ribbons.Ribbon1.group3.Visible = false;
                Properties.Settings.Default.g2 = 0;
                Properties.Settings.Default.Save();
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked)
            {
                Globals.Ribbons.Ribbon1.group4.Visible = true;
                Properties.Settings.Default.g3 = 1;
                Properties.Settings.Default.Save();
            }
            else
            {
                Globals.Ribbons.Ribbon1.group4.Visible = false;
                Properties.Settings.Default.g3 = 0;
                Properties.Settings.Default.Save();
            }
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                Globals.Ribbons.Ribbon1.group1.Visible = true;
                Properties.Settings.Default.g4 = 1;
                Properties.Settings.Default.Save();
            }
            else
            {
                Globals.Ribbons.Ribbon1.group1.Visible = false;
                Properties.Settings.Default.g4 = 0;
                Properties.Settings.Default.Save();
            }
        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked)
            {
                Globals.Ribbons.Ribbon1.group7.Visible = true;
                Properties.Settings.Default.g5 = 1;
                Properties.Settings.Default.Save();
            }
            else
            {
                Globals.Ribbons.Ribbon1.group7.Visible = false;
                Properties.Settings.Default.g5 = 0;
                Properties.Settings.Default.Save();
            }
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked)
            {
                Globals.Ribbons.Ribbon1.group6.Visible = true;
                Properties.Settings.Default.g6 = 1;
                Properties.Settings.Default.Save();
                checkBox7.Enabled = true;
            }
            else
            {
                Globals.Ribbons.Ribbon1.group6.Visible = false;
                Properties.Settings.Default.g6 = 0;
                Properties.Settings.Default.Save();
                checkBox7.Enabled = false;
            }
        }

        private void setting1_Load(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.g1 == 0)
            {
                checkBox1.Checked = false;
            }
            else
            {
                checkBox1.Checked = true;
            }
            if (Properties.Settings.Default.g2 == 0)
            {
                checkBox2.Checked = false;
            }
            else
            {
                checkBox2.Checked = true;
            }
            if (Properties.Settings.Default.g3 == 0)
            {
                checkBox3.Checked = false;
            }
            else
            {
                checkBox3.Checked = true;
            }
            if (Properties.Settings.Default.g4 == 0)
            {
                checkBox4.Checked = false;
            }
            else
            {
                checkBox4.Checked = true;
            }
            if (Properties.Settings.Default.g5 == 0)
            {
                checkBox5.Checked = false;
            }
            else
            {
                checkBox5.Checked = true;
            }
            if (Properties.Settings.Default.g6 == 0)
            {
                checkBox6.Checked = false;
            }
            else
            {
                checkBox6.Checked = true;
            }
            if (Properties.Settings.Default.g7 == 0)
            {
                checkBox8.Checked = false;
            }
            else
            {
                checkBox8.Checked = true;
            }
            if (Properties.Settings.Default.morph == 0)
            {
                checkBox7.Checked = false;
            }
            else
            {
                checkBox7.Checked = true;
            }
            if (Properties.Settings.Default.tab2 == 0)
            {
                checkBox9.Checked = false;
            }
            else
            {
                checkBox9.Checked = true;
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked)
            {
                Globals.Ribbons.Ribbon1.menu26.Visible = true;
                Properties.Settings.Default.morph = 1;
                Properties.Settings.Default.Save();
            }
            else
            {
                Globals.Ribbons.Ribbon1.menu26.Visible = false;
                Properties.Settings.Default.morph = 0;
                Properties.Settings.Default.Save();
            }
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox8.Checked)
            {
                Globals.Ribbons.Ribbon1.group8.Visible = true;
                Properties.Settings.Default.g7 = 1;
                Properties.Settings.Default.Save();
            }
            else
            {
                Globals.Ribbons.Ribbon1.group8.Visible = false;
                Properties.Settings.Default.g7 = 0;
                Properties.Settings.Default.Save();
            }
        }

        private void checkBox9_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox9.Checked)
            {
                Globals.Ribbons.Ribbon1.Tabs[1].Visible = true;
                Properties.Settings.Default.tab2 = 1;
                Properties.Settings.Default.Save();
            }
            else
            {
                if (Properties.Settings.Default.tab1 == 0)
                {
                    Globals.Ribbons.Ribbon1.Tabs[0].Visible = true;
                    Properties.Settings.Default.tab1 = 1;
                    Properties.Settings.Default.Save();
                    Globals.Ribbons.Ribbon1.button247.Label = "隐藏主卡";
                }
                Globals.Ribbons.Ribbon1.Tabs[1].Visible = false;
                Properties.Settings.Default.tab2 = 0;
                Properties.Settings.Default.Save();
            }
        }
    }
}
