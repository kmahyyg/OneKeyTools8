using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using forms = System.Windows.Forms;

namespace OneKeyTools
{
    public partial class Picture_inPic : Form
    {
        public Picture_inPic()
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

        private void button2_Click(object sender, EventArgs e)
        {
            Picture_inPic.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button171.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.ShapeRange range = sel.ShapeRange;
                int count = range.Count;
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape pic = range[i];
                    if (pic.Type == Office.MsoShapeType.msoPicture)
                    {
                        PowerPoint.Shape npic = pic.Duplicate()[1];
                        float pw = app.ActivePresentation.PageSetup.SlideWidth;
                        float ph = app.ActivePresentation.PageSetup.SlideHeight;
                        if (checkBox1.Checked)
                        {
                            if (npic.LockAspectRatio == Office.MsoTriState.msoTrue)
                            {
                                npic.LockAspectRatio = Office.MsoTriState.msoFalse;
                            }
                            npic.Width = pw;
                            npic.Height = ph;
                            npic.Left = 0;
                            npic.Top = 0;
                            pic.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                        }
                        else
                        {
                            npic.Width = pic.Width * 1.5f;
                            npic.Height = pic.Height * 1.5f;
                            pic.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                        }
                        if (checkBox2.Checked)
                        {
                            if (npic.Fill.PictureEffects.Count == 0)
	                        {
                                Office.PictureEffect piceff = npic.Fill.PictureEffects.Insert(Office.MsoPictureEffectType.msoEffectBlur, 0);
                                piceff.EffectParameters[1].Value = 30f;
	                        }
                            else
                            {
                                int en = -1;
                                for (int j = 1; j <= npic.Fill.PictureEffects.Count; j++)
                                {
                                    if (npic.Fill.PictureEffects[j].Type == Office.MsoPictureEffectType.msoEffectBlur)
                                    {
                                        Office.PictureEffect piceff = npic.Fill.PictureEffects[j];
                                        piceff.EffectParameters[1].Value = 30f;
                                        en = 1;
                                    }
                                }
                                if (en == -1)
                                {
                                    Office.PictureEffect piceff = npic.Fill.PictureEffects.Insert(Office.MsoPictureEffectType.msoEffectBlur, 0);
                                    piceff.EffectParameters[1].Value = 30f;
                                }
                            }
                        }
                        if (checkBox3.Checked)
                        {
                            if (npic.Fill.PictureEffects.Count == 0)
                            {
                                Office.PictureEffect piceff = npic.Fill.PictureEffects.Insert(Office.MsoPictureEffectType.msoEffectBrightnessContrast, 0);
                                piceff.EffectParameters[1].Value = -0.43f;
                            }
                            else
                            {
                                int en = -1;
                                for (int j = 1; j <= npic.Fill.PictureEffects.Count; j++)
                                {
                                    if (npic.Fill.PictureEffects[j].Type == Office.MsoPictureEffectType.msoEffectBrightnessContrast)
                                    {
                                        Office.PictureEffect piceff = npic.Fill.PictureEffects[j];
                                        piceff.EffectParameters[1].Value = -0.43f;
                                        en = 1;
                                    }
                                }
                                if (en == -1)
                                {
                                    Office.PictureEffect piceff = npic.Fill.PictureEffects.Insert(Office.MsoPictureEffectType.msoEffectBrightnessContrast, 0);
                                    piceff.EffectParameters[1].Value = -0.43f;
                                }
                            }
                        }
                        List<string> name = new List<string>();
                        name.Add(npic.Name);
                        name.Add(pic.Name);
                        if (radioButton1.Checked)
                        { 
                            if (checkBox1.Checked)
                            {

                                slide.Shapes.Range(name.ToArray()).Align(Office.MsoAlignCmd.msoAlignMiddles, Office.MsoTriState.msoFalse);
                                slide.Shapes.Range(name.ToArray()).Align(Office.MsoAlignCmd.msoAlignCenters, Office.MsoTriState.msoFalse);
                            }
                            else
                            {
                                npic.Left = pic.Left + pic.Width / 2 - npic.Width / 2;
                                npic.Top = pic.Top + pic.Height / 2 - npic.Height / 2;
                            }
                        }
                        else if (radioButton2.Checked)
                        {
                            if (checkBox1.Checked)
                            {
                                slide.Shapes.Range(name.ToArray()).Align(Office.MsoAlignCmd.msoAlignRights, Office.MsoTriState.msoFalse);
                                slide.Shapes.Range(name.ToArray()).Align(Office.MsoAlignCmd.msoAlignBottoms, Office.MsoTriState.msoFalse);
                            }
                            else
                            {
                                npic.Left = pic.Left + pic.Width - npic.Width;
                                npic.Top = pic.Top + pic.Height - npic.Height;
                            }
                        }
                        slide.Shapes.Range(name.ToArray()).Group();
                    }
                }
            }
            else
            {
                MessageBox.Show("请选中图片");
            }
        }

    }
}
