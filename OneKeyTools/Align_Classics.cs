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
    public partial class Align_Classics : Form
    {
        public Align_Classics()
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

        private void button13_Click(object sender, EventArgs e)
        {
            Align_Classics.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button202.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                if (range.Count == 1)
                {
                    range[1].Left = 0;
                }
                else
                {
                    for (int i = 2; i <= range.Count; i++)
                    {
                        range[i].Left = range[1].Left;
                    }
                }  
            }
            else
            {
                PowerPoint.SlideRange srange = sel.SlideRange;
                foreach (PowerPoint.Slide item in srange)
                {
                    for (int i = 1; i <= item.Shapes.Count; i++)
                    {
                        item.Shapes[i].Left = 0;
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                if (range.Count == 1)
                {
                    range[1].Left = app.ActivePresentation.PageSetup.SlideWidth - range[1].Width;
                }
                else
                {
                    for (int i = 2; i <= range.Count; i++)
                    {
                        range[i].Left = range[1].Left + range[1].Width - range[i].Width;
                    }
                }
            }
            else
            {
                PowerPoint.SlideRange srange = sel.SlideRange;
                float swidth = app.ActivePresentation.PageSetup.SlideWidth;
                foreach (PowerPoint.Slide item in srange)
                {
                    for (int i = 1; i <= item.Shapes.Count; i++)
                    {
                        item.Shapes[i].Left = swidth - item.Shapes[i].Width;
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                if (range.Count == 1)
                {
                    range[1].Top = 0;
                }
                else
                {
                    for (int i = 2; i <= range.Count; i++)
                    {
                        range[i].Top = range[1].Top;
                    }
                }
            }
            else
            {
                PowerPoint.SlideRange srange = sel.SlideRange;
                foreach (PowerPoint.Slide item in srange)
                {
                    for (int i = 1; i <= item.Shapes.Count; i++)
                    {
                        item.Shapes[i].Top = 0;
                    }
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                if (range.Count == 1)
                {
                    range[1].Top = app.ActivePresentation.PageSetup.SlideHeight - range[1].Height;
                }
                else
                {
                    for (int i = 2; i <= range.Count; i++)
                    {
                        range[i].Top = range[1].Top + range[1].Height - range[i].Height;
                    }
                }
            }
            else
            {
                PowerPoint.SlideRange srange = sel.SlideRange;
                float sheight = app.ActivePresentation.PageSetup.SlideHeight;
                foreach (PowerPoint.Slide item in srange)
                {
                    for (int i = 1; i <= item.Shapes.Count; i++)
                    {
                        item.Shapes[i].Top = sheight - item.Shapes[i].Height;
                    }
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                if (range.Count == 1)
                {
                    range[1].Left = app.ActivePresentation.PageSetup.SlideWidth / 2 - range[1].Width / 2;
                }
                else
                {
                    for (int i = 2; i <= range.Count; i++)
                    {
                        range[i].Left = range[1].Left + range[1].Width / 2 - range[i].Width / 2;
                    }
                }
            }
            else
            {
                PowerPoint.SlideRange srange = sel.SlideRange;
                float swidth = app.ActivePresentation.PageSetup.SlideWidth;
                foreach (PowerPoint.Slide item in srange)
                {
                    for (int i = 1; i <= item.Shapes.Count; i++)
                    {
                        item.Shapes[i].Left = swidth / 2 - item.Shapes[i].Width / 2;
                    }
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.ShapeRange range = sel.ShapeRange;
                if (sel.HasChildShapeRange)
                {
                    range = sel.ChildShapeRange;
                }
                if (range.Count == 1)
                {
                    range[1].Top = app.ActivePresentation.PageSetup.SlideHeight / 2 - range[1].Height / 2;
                }
                else
                {
                    for (int i = 2; i <= range.Count; i++)
                    {
                        range[i].Top = range[1].Top + range[1].Height / 2 - range[i].Height / 2;
                    }
                }
            }
            else
            {
                PowerPoint.SlideRange srange = sel.SlideRange;
                float sheight = app.ActivePresentation.PageSetup.SlideHeight;
                foreach (PowerPoint.Slide item in srange)
                {
                    for (int i = 1; i <= item.Shapes.Count; i++)
                    {
                        item.Shapes[i].Top = sheight / 2 - item.Shapes[i].Height / 2;
                    }
                }
            }
        }



    }
}
