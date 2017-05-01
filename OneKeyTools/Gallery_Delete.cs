using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace OneKeyTools
{
    public partial class Gallery_Delete : Form
    {
        public Gallery_Delete()
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

        string fName = "";

        private void 刷新ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Location = "";
            RegistryKey path = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Slibe\OneKeyTools", false);
            if (Properties.Settings.Default.Galleryfolder == "")
            {
                Location = path.GetValue("Path", "").ToString();
            }
            else
            {
                Location = Properties.Settings.Default.Galleryfolder;
            }
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = Location;
            openFileDialog.Filter = "PowerPoint演示文稿|*.pptx";
            openFileDialog.FilterIndex = 1;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                fName = openFileDialog.FileName;
                if (fName.Contains("pptx"))
                {
                    PowerPoint.Application pptapp = new PowerPoint.Application();
                    PowerPoint.Presentation pptpr = pptapp.Presentations.Open(fName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse);

                    listView1.Clear();
                    this.listView1.Refresh();
                    imageList1.Images.Clear();

                    if (!Directory.Exists(Location + @"temp_gallery2\"))
                    {
                        Directory.CreateDirectory(Location + @"temp_gallery2\");
                    }
                    System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(Location + @"temp_gallery2\");

                    int n = 0;
                    foreach (PowerPoint.Slide oslide in pptpr.Slides)
                    {
                        if (oslide.Shapes.Count != 0)
                        {
                            for (int i = 1; i <= oslide.Shapes.Count; i++)
                            {
                                PowerPoint.Shape oshape = oslide.Shapes[i];
                                int k = dir.GetFiles().Length + 1;
                                string npath = Location + @"temp_gallery2\gshape_" + k + ".png";
                                oshape.Export(npath, PowerPoint.PpShapeFormat.ppShapeFormatPNG, 300, 300);
                                Image img = Image.FromFile(npath);
                                Bitmap bmp = new Bitmap(img);
                                img.Dispose();
                                imageList1.Images.Add(bmp);
                            }
                            n += 1;
                        }
                    }
                    Directory.Delete(Location + @"temp_gallery2\", true);
                    pptpr.Close();

                    if (n == 0)
                    {
                        MessageBox.Show("所选文稿中没有图形，请重新加载");
                    }
                    else
                    {
                        listView1.View = View.LargeIcon;
                        listView1.LargeImageList = imageList1;
                        listView1.BeginUpdate();
                        for (int i = 0; i < imageList1.Images.Count; i++)
                        {
                            ListViewItem item = new ListViewItem();
                            item.ImageIndex = i;
                            listView1.Items.Add(item);
                        }
                        this.listView1.EndUpdate();
                    }
                }
            }
        }

        private void 删除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (fName != "")
            {
                List<int> del = new List<int>();
                int n = listView1.SelectedItems.Count;
                for (int i = n - 1; i >= 0; i--)
                {
                    ListViewItem item = listView1.SelectedItems[i];
                    item.Remove();
                    del.Add(item.ImageIndex);
                }

                listView1.Clear();
                this.listView1.Refresh();
                imageList1.Images.Clear();

                if (!Directory.Exists(Location + @"temp_gallery2\"))
                {
                    Directory.CreateDirectory(Location + @"temp_gallery2\");
                }

                System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(Location + @"temp_gallery2\");

                PowerPoint.Application pptapp = new PowerPoint.Application();
                PowerPoint.Presentation pptpr = pptapp.Presentations.Open(fName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse);
                for (int k = del.Count() - 1; k >= 0; k--)
                {
                    int cn = del[k];
                    if (k != del.Count() - 1)
                    {
                        cn = cn - 1;
                    }
                    int sn = 0;
                    if (pptpr.Slides.Count > 1)
                    {
                        for (int i = 1; i <= pptpr.Slides.Count; i++)
                        {
                            PowerPoint.Slide oslide = pptpr.Slides[i];
                            if (oslide.Shapes.Count != 0)
                            {
                                sn += oslide.Shapes.Count;
                                if (cn + 1 <= sn)
                                {
                                    oslide.Shapes[oslide.Shapes.Count - (sn - cn - 1)].Delete();
                                    break;
                                }
                            }
                        }
                    }
                    else
                    {
                        pptpr.Slides[1].Shapes[cn + 1].Delete();
                    }

                }
                pptpr.Save();
                if (pptpr.Slides.Count != 0)
                {
                    for (int i = 1; i <= pptpr.Slides.Count; i++)
                    {
                        PowerPoint.Slide oslide = pptpr.Slides[i];
                        if (oslide.Shapes.Count != 0)
                        {
                            for (int j = 1; j <= oslide.Shapes.Count; j++)
                            {
                                PowerPoint.Shape oshape = oslide.Shapes[j];
                                int k = dir.GetFiles().Length + 1;
                                oshape.Export(Location + @"temp_gallery2\gshape_" + k + ".png", PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                                Image img = Image.FromFile(Location + @"temp_gallery2\gshape_" + k + ".png");
                                Bitmap bmp = new Bitmap(img);
                                img.Dispose();
                                imageList1.Images.Add(bmp);
                            }
                        }
                    }
                }

                Properties.Settings.Default.GalleryRefresh = 1;
                Directory.Delete(Location + @"temp_gallery2\", true);

                pptpr.Close();
                listView1.LargeImageList = imageList1;
                listView1.BeginUpdate();
                for (int i = 0; i < imageList1.Images.Count; i++)
                {
                    ListViewItem item = new ListViewItem();
                    item.ImageIndex = i;
                    listView1.Items.Add(item);
                }
                this.listView1.EndUpdate();
            }
            else
            {
                MessageBox.Show("请先加载库");
            }
        }

        private void 删除所有ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (fName != "")
            {
                PowerPoint.Application pptapp = new PowerPoint.Application();
                PowerPoint.Presentation pptpr = pptapp.Presentations.Open(fName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse);
                int n = 0;
                for (int i = pptpr.Slides.Count; i >= 1; i--)
                {
                    PowerPoint.Slide oslide = pptpr.Slides[i];
                    n += oslide.Shapes.Count;
                    oslide.Delete();
                }
                pptpr.Slides.AddSlide(1, pptpr.SlideMaster.CustomLayouts[7]);
                pptpr.Save();
                Properties.Settings.Default.GalleryRefresh = 1;
                pptpr.Close();
                listView1.Items.Clear();
                imageList1.Images.Clear();
                MessageBox.Show("共删除了 " + n + " 个图形");
            }
            else
            {
                MessageBox.Show("请先加载库");
            }
        }

        private void 删除空白页ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (fName != "")
            {
                PowerPoint.Application pptapp = new PowerPoint.Application();
                PowerPoint.Presentation pptpr = pptapp.Presentations.Open(fName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse);
                int n = 0;
                for (int i = pptpr.Slides.Count; i >= 1; i--)
                {
                    PowerPoint.Slide oslide = pptpr.Slides[i];
                    if (oslide.Shapes.Count == 0)
                    {
                        oslide.Delete();
                        n += 1;
                    }
                }
                if (pptpr.Slides.Count == 0)
                {
                    pptpr.Slides.AddSlide(1, pptpr.SlideMaster.CustomLayouts[7]);
                }
                pptpr.Save();
                Properties.Settings.Default.GalleryRefresh = 1;
                pptpr.Close();
                MessageBox.Show("共删除了 " + n + " 个空白页");
            }
            else
            {
                MessageBox.Show("请先加载库");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Gallery_Delete.ActiveForm.Close();
        }


    }
}
