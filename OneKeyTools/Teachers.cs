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
    public partial class Teachers : Form
    {
        public Teachers()
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
            Teachers.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button27.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            app.SlideShowWindows[1].View.PointerType = PowerPoint.PpSlideShowPointerType.ppSlideShowPointerArrow;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            app.SlideShowWindows[1].View.PointerType = PowerPoint.PpSlideShowPointerType.ppSlideShowPointerPen;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            app.SlideShowWindows[1].View.PointerType = PowerPoint.PpSlideShowPointerType.ppSlideShowPointerEraser;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            app.SlideShowWindows[1].View.State = PowerPoint.PpSlideShowState.ppSlideShowBlackScreen;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            app.SlideShowWindows[1].View.State = PowerPoint.PpSlideShowState.ppSlideShowWhiteScreen;

        }

        private void button8_Click(object sender, EventArgs e)
        {
            app.SlideShowWindows[1].View.Exit();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            MessageBox.Show("本工具为  @课堂大厨  定制功能");
            System.Diagnostics.Process.Start("http://weibo.com/classchef"); 
        }

    }
}
