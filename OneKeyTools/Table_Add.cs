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
    public partial class Table_Add : Form
    {
        public Table_Add()
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
            Table_Add.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button249.Enabled = true;
        }

        private void table_add1_Load(object sender, EventArgs e)
        {
            textBox1.Text = Properties.Settings.Default.table_addr.ToString();
            textBox2.Text = Properties.Settings.Default.table_addc.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.table_addr = int.Parse(textBox1.Text.Trim());
            Properties.Settings.Default.table_addc = int.Parse(textBox2.Text.Trim());
            Properties.Settings.Default.Save();
            MessageBox.Show("保存成功");
        }


    }
}
