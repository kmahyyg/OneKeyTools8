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
    public partial class Picture_Assist : Form
    {
        public Picture_Assist()
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

        private float PtoCM(float p)
        {
            float cm = (float)(p * 2.54 / 72);
            return cm;
        }

        private float CMtoP(float cm)
        {
            float p = (float)(cm * 72 / 2.54);
            return p;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && ((!sel.HasChildShapeRange && (sel.ShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture)) || (sel.HasChildShapeRange && (sel.ChildShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ChildShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture))))
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
                float n = float.Parse(textBox4.Text.Trim());
                n = CMtoP(n);
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape pic = range[i];
                    pic.PictureFormat.Crop.PictureOffsetY -= n;
                }
            }
            else
            {
                MessageBox.Show("请先选择一张图片");
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && ((!sel.HasChildShapeRange && (sel.ShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture)) || (sel.HasChildShapeRange && (sel.ChildShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ChildShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture))))
            {
                string n = textBox2.Text.Trim();
                float b;
                if (n == "" || !float.TryParse(n, out b) || n == "0")
                {
                    MessageBox.Show("请输入大于0的数值");
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
                    for (int i = 1; i <= count; i++)
                    {
                        PowerPoint.Shape pic = range[i];
                        pic.PictureFormat.Crop.PictureHeight = CMtoP(float.Parse(n));
                    }
                }
            }
            else
            {
                MessageBox.Show("请先选择一张图片");
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && ((!sel.HasChildShapeRange && (sel.ShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture)) || (sel.HasChildShapeRange && (sel.ChildShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ChildShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture))))
            {
                string n = textBox3.Text.Trim();
                float b;
                if (n == "" || !float.TryParse(n, out b) || n == "0")
                {
                    MessageBox.Show("请输入大于0的数值");
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
                    for (int i = 1; i <= count; i++)
                    {
                        PowerPoint.Shape pic = range[i];
                        pic.PictureFormat.Crop.PictureWidth = CMtoP(float.Parse(n));
                    }
                }
            }
            else
            {
                MessageBox.Show("请先选择一张图片");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && ((!sel.HasChildShapeRange && (sel.ShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture)) || (sel.HasChildShapeRange && (sel.ChildShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ChildShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture))))
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
                float n = float.Parse(textBox4.Text.Trim());
                n = CMtoP(n);
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape pic = range[i];
                    pic.PictureFormat.Crop.PictureOffsetY += n;
                }
            }
            else
            {
                MessageBox.Show("请先选择一张图片");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && ((!sel.HasChildShapeRange && (sel.ShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture)) || (sel.HasChildShapeRange && (sel.ChildShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ChildShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture))))
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
                float n = float.Parse(textBox4.Text.Trim());
                n = CMtoP(n);
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape pic = range[i];
                    pic.PictureFormat.Crop.PictureOffsetX -= n;
                }
            }
            else
            {
                MessageBox.Show("请先选择一张图片");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && ((!sel.HasChildShapeRange && (sel.ShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture)) || (sel.HasChildShapeRange && (sel.ChildShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ChildShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture))))
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
                float n = float.Parse(textBox4.Text.Trim());
                n = CMtoP(n);
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape pic = range[i];
                    pic.PictureFormat.Crop.PictureOffsetX += n;
                }
            }
            else
            {
                MessageBox.Show("请先选择一张图片");
            }
        }

        float h0 = 0;
        float w0 = 0;
        int id = 0;
        private void button5_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && ((!sel.HasChildShapeRange && (sel.ShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture)) || (sel.HasChildShapeRange && (sel.ChildShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ChildShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture))))
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
                PowerPoint.Shape pic = range[1];           
                if (h0==0)
                {
                    h0 = pic.PictureFormat.Crop.PictureHeight;
                    w0 = pic.PictureFormat.Crop.PictureWidth;
                    id = pic.Id;
                    textBox2.Text = PtoCM(pic.PictureFormat.Crop.PictureHeight).ToString();
                    textBox3.Text = PtoCM(pic.PictureFormat.Crop.PictureWidth).ToString();
                }
                else
                {
                    if (pic.Id == id)
                    {
                        pic.PictureFormat.Crop.PictureHeight = h0;
                        pic.PictureFormat.Crop.PictureWidth = w0;
                        textBox2.Text = PtoCM(pic.PictureFormat.Crop.PictureHeight).ToString();
                        textBox3.Text = PtoCM(pic.PictureFormat.Crop.PictureWidth).ToString();  
                    }
                    else
                    {
                        h0 = pic.PictureFormat.Crop.PictureHeight;
                        w0 = pic.PictureFormat.Crop.PictureWidth;
                        id = pic.Id;
                        textBox2.Text = PtoCM(pic.PictureFormat.Crop.PictureHeight).ToString();
                        textBox3.Text = PtoCM(pic.PictureFormat.Crop.PictureWidth).ToString();  
                    }
                    
                }
            }
            else
            {
                MessageBox.Show("请先选择一张图片");
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Picture_Assist.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button67.Enabled = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && ((!sel.HasChildShapeRange && (sel.ShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture)) || (sel.HasChildShapeRange && (sel.ChildShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ChildShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture))))
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
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape pic = range[i];
                    pic.PictureFormat.Crop.PictureOffsetX = (pic.PictureFormat.Crop.ShapeWidth - pic.PictureFormat.Crop.PictureWidth)/1024;
                    pic.PictureFormat.Crop.PictureOffsetY = (pic.PictureFormat.Crop.ShapeHeight - pic.PictureFormat.Crop.PictureHeight)/1024;
                }
            }
            else
            {
                MessageBox.Show("请先选择一张图片");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && ((!sel.HasChildShapeRange && (sel.ShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture)) || (sel.HasChildShapeRange && (sel.ChildShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ChildShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture))))
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
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape pic = range[i];
                    pic.PictureFormat.Crop.PictureHeight = pic.PictureFormat.Crop.ShapeHeight;
                    pic.PictureFormat.Crop.PictureWidth = pic.PictureFormat.Crop.ShapeWidth;
                }
            }
            else
            {
                MessageBox.Show("请先选择一张图片");
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && ((!sel.HasChildShapeRange && (sel.ShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture)) || (sel.HasChildShapeRange && (sel.ChildShapeRange.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || sel.ChildShapeRange.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPicture))))
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
                float n = float.Parse(textBox1.Text.Trim()) / 100;
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape pic = range[i];
                    pic.PictureFormat.Crop.PictureHeight = pic.PictureFormat.Crop.PictureHeight * n;
                    pic.PictureFormat.Crop.PictureWidth = pic.PictureFormat.Crop.PictureWidth * n;
                }
            }
            else
            {
                MessageBox.Show("请先选择一张图片");
            }
        }

    }
}
