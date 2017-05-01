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
using Microsoft.Win32;

namespace OneKeyTools
{
    public partial class Picture_Settings : Form
    {
        public Picture_Settings()
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

        private void dpi1_Load(object sender, EventArgs e)
        {
            textBox1.Text = Properties.Settings.Default.dpi.ToString();
            textBox2.Text = Properties.Settings.Default.pwidth.ToString();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            float dpi = float.Parse(textBox1.Text.Trim());
            if (checkBox1.Checked)
            {
                if (dpi >= 300)
                {
                    textBox2.Text = "4000";
                }
                else if (dpi >=250)
                {
                    textBox2.Text = "3333";
                }
                else if (dpi >= 200)
                {
                    textBox2.Text = "2667";
                }
                else if (dpi >= 150)
                {
                    textBox2.Text = "2000";
                }
                else if (dpi >= 100)
                {
                    textBox2.Text = "1333";
                }
                else if (dpi >= 96)
                {
                    textBox2.Text = "1280";
                }
                else
                {
                    textBox2.Text = "720";
                }
            }
            else
            {
                if (dpi >= 300)
                {
                    textBox2.Text = "3000";
                }
                else if (dpi >= 250)
                {
                    textBox2.Text = "2500";
                }
                else if (dpi >= 200)
                {
                    textBox2.Text = "2000";
                }
                else if (dpi >= 150)
                {
                    textBox2.Text = "1500";
                }
                else if (dpi >= 100)
                {
                    textBox2.Text = "1000";
                }
                else if (dpi >= 96)
                {
                    textBox2.Text = "960";
                }
                else
                {
                    textBox2.Text = "720";
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            float dpi = float.Parse(textBox1.Text.Trim());
            int pw = int.Parse(textBox2.Text.Trim());
            Properties.Settings.Default.dpi = dpi;
            Properties.Settings.Default.pwidth = pw;
            Properties.Settings.Default.Save();
            MessageBox.Show("已修改");
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            float dpi = float.Parse(textBox1.Text.Trim());
            if (checkBox1.Checked)
            {
                if (dpi >= 300)
                {
                    textBox2.Text = "4000";
                }
                else if (dpi >= 250)
                {
                    textBox2.Text = "3333";
                }
                else if (dpi >= 200)
                {
                    textBox2.Text = "2667";
                }
                else if (dpi >= 150)
                {
                    textBox2.Text = "2000";
                }
                else if (dpi >= 100)
                {
                    textBox2.Text = "1333";
                }
                else if (dpi >= 96)
                {
                    textBox2.Text = "1280";
                }
                else
                {
                    textBox2.Text = "667";
                }
            }
            else
            {
                if (dpi >= 300)
                {
                    textBox2.Text = "3000";
                }
                else if (dpi >= 250)
                {
                    textBox2.Text = "2500";
                }
                else if (dpi >= 200)
                {
                    textBox2.Text = "2000";
                }
                else if (dpi >= 150)
                {
                    textBox2.Text = "1500";
                }
                else if (dpi >= 100)
                {
                    textBox2.Text = "1000";
                }
                else if (dpi >= 96)
                {
                    textBox2.Text = "960";
                }
                else
                {
                    textBox2.Text = "500";
                }
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            //float pw = float.Parse(textBox2.Text.Trim());
            //if (checkBox1.Checked)
            //{
            //    if (pw >= 4000)
            //    {
            //        textBox1.Text = "300";
            //    }
            //    else if (pw >= 3333)
            //    {
            //        textBox1.Text = "250";
            //    }
            //    else if (pw >= 2667)
            //    {
            //        textBox1.Text = "200";
            //    }
            //    else if (pw >= 2000)
            //    {
            //        textBox1.Text = "150";
            //    }
            //    else if (pw >= 1333)
            //    {
            //        textBox1.Text = "100";
            //    }
            //    else if (pw >= 1280)
            //    {
            //        textBox1.Text = "96";
            //    }
            //    else
            //    {
            //        textBox1.Text = "50";
            //    }
            //}
            //else
            //{
            //    if (pw >= 3000)
            //    {
            //        textBox1.Text = "300";
            //    }
            //    else if (pw >= 2500)
            //    {
            //        textBox1.Text = "250";
            //    }
            //    else if (pw >= 2000)
            //    {
            //        textBox1.Text = "200";
            //    }
            //    else if (pw >= 1500)
            //    {
            //        textBox1.Text = "150";
            //    }
            //    else if (pw >= 1000)
            //    {
            //        textBox1.Text = "100";
            //    }
            //    else if (pw >= 960)
            //    {
            //        textBox1.Text = "96";
            //    }
            //    else
            //    {
            //        textBox1.Text = "50";
            //    }
            //}
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = "96";
            textBox2.Text = "720";
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text=="初始化")
            {
                RegistryKey path2007 = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Office\12.0\PowerPoint\Options", true);
                if (path2007 != null)
                {
                    string dpi = path2007.GetValue("ExportBitmapResolution", "00001").ToString();
                    if (dpi == "00001" || dpi != "300")
                    {
                        path2007.SetValue("ExportBitmapResolution", 300, RegistryValueKind.DWord);
                    }
                }

                RegistryKey path2010 = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Office\14.0\PowerPoint\Options", true);
                if (path2010 != null)
                {
                    string dpi = path2010.GetValue("ExportBitmapResolution", "00001").ToString();
                    if (dpi == "00001" || dpi != "300")
                    {
                        path2010.SetValue("ExportBitmapResolution", 300, RegistryValueKind.DWord);
                    }
                }

                RegistryKey path2013 = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Office\15.0\PowerPoint\Options", true);
                if (path2013 != null)
                {
                    string dpi = path2013.GetValue("ExportBitmapResolution", "00001").ToString();
                    if (dpi == "00001" || dpi != "300")
                    {
                        path2013.SetValue("ExportBitmapResolution", 300, RegistryValueKind.DWord);
                    }
                }

                RegistryKey path2016 = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Office\16.0\PowerPoint\Options", true);
                if (path2016 != null)
                {
                    string dpi = path2016.GetValue("ExportBitmapResolution", "00001").ToString();
                    if (dpi == "00001" || dpi != "300")
                    {
                        path2016.SetValue("ExportBitmapResolution", 300, RegistryValueKind.DWord);
                    }
                }

                MessageBox.Show("已把 ExportBitmapResolution键 添加到注册表，以后使用无需重复本功能");
            }
            else if (comboBox1 .Text =="恢复默认")
            {
                RegistryKey path2007 = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Office\12.0\PowerPoint\Options", true);
                if (path2007 != null)
                {
                    path2007.DeleteValue("ExportBitmapResolution", false);
                }

                RegistryKey path2010 = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Office\14.0\PowerPoint\Options", true);
                if (path2010 != null)
                {
                    path2010.DeleteValue("ExportBitmapResolution", false);
                }

                RegistryKey path2013 = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Office\15.0\PowerPoint\Options", true);
                if (path2013 != null)
                {
                    path2013.DeleteValue("ExportBitmapResolution", false);
                }

                RegistryKey path2016 = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Office\16.0\PowerPoint\Options", true);
                if (path2016 != null)
                {
                    path2016.DeleteValue("ExportBitmapResolution", false);
                }
                MessageBox.Show("已把 ExportBitmapResolution键 从注册表删除，以后要导出300DPI图请重新初始化");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Picture_Settings.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button126.Enabled = true;
        }
    }
}
