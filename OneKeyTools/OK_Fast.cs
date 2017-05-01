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
using System.Drawing.Imaging;
using System.Diagnostics;

namespace OneKeyTools
{
    public partial class OK_Fast : Form
    {
        public OK_Fast()
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

        private int Hsl2Rgb(int h, int s, int l)
        {
            float nh = (float)h / 255;
            float ns = (float)s / 255;
            float nl = (float)l / 255;
            float nr = 0;
            float ng = 0;
            float nb = 0;

            if (ns == 0)
            {
                nr = nl;
                ng = nl;
                nb = nl;
            }
            else
            {
                float q = 0;
                if (nl < 0.5f)
                {
                    q = nl * (1 + ns);
                }
                else
                {
                    q = nl + ns - nl * ns;
                }

                float p = 2 * nl - q;
                float tr = nh + (float)1 / 3;
                float tg = nh;
                float tb = nh - (float)1 / 3;

                if (tr < 0)
                {
                    tr = tr + 1;
                }
                else
                {
                    if (tr > 1)
                    {
                        tr = tr - 1;
                    }
                    else
                    {
                        tr = tr + 0;
                    }
                }

                if (tg < 0)
                {
                    tg = tg + 1;
                }
                else
                {
                    if (tg > 1)
                    {
                        tg = tg - 1;
                    }
                    else
                    {
                        tg = tg + 0;
                    }
                }

                if (tb < 0)
                {
                    tb = tb + 1;
                }
                else
                {
                    if (tb > 1)
                    {
                        tb = tb - 1;
                    }
                    else
                    {
                        tb = tb + 0;
                    }
                }

                if (tr < (float)1 / 6)
                {
                    nr = p + (q - p) * 6 * tr;
                }
                else
                {
                    if (tr < (float)1 / 2)
                    {
                        nr = q;
                    }
                    else
                    {
                        if (tr < (float)2 / 3)
                        {
                            nr = p + (q - p) * 6 * ((float)2 / 3 - tr);
                        }
                        else
                        {
                            nr = p;
                        }
                    }
                }

                if (tg < (float)1 / 6)
                {
                    ng = p + (q - p) * 6 * tg;
                }
                else
                {
                    if (tg < (float)1 / 2)
                    {
                        ng = q;
                    }
                    else
                    {
                        if (tg < (float)2 / 3)
                        {
                            ng = p + (q - p) * 6 * ((float)2 / 3 - tg);
                        }
                        else
                        {
                            ng = p;
                        }
                    }
                }

                if (tb < (float)1 / 6)
                {
                    nb = p + (q - p) * 6 * tb;
                }
                else
                {
                    if (tb < (float)1 / 2)
                    {
                        nb = q;
                    }
                    else
                    {
                        if (tb < (float)2 / 3)
                        {
                            nb = p + (q - p) * 6 * ((float)2 / 3 - tb);
                        }
                        else
                        {
                            nb = p;
                        }
                    }
                }
            }
            int r = (int)(nr * 255);
            int g = (int)(ng * 255);
            int b = (int)(nb * 255);
            int rgb = r + g * 256 + b * 256 * 256;
            return rgb;
        }

        private int Rgb2Hsl(int r, int g, int b)
        {
            float h = 0;
            float s = 0;
            float l = 0;
            float max = Math.Max(Math.Max(r, g), b);
            float min = Math.Min(Math.Min(r, g), b);

            if (max == min)
            {
                h = 0;
            }
            else
            {
                if (max == r)
                {
                    if (g >= b)
                    {
                        h = 255 / 6 * (g - b) / (max - min) + 0;
                    }
                    else
                    {
                        h = 255 / 6 * (g - b) / (max - min) + 255;
                    }
                }
                if (max == g & max != r)
                {
                    h = 255 / 6 * (b - r) / (max - min) + 255 / 3;
                }
                if (max == b && max != g)
                {
                    h = 255 / 6 * (r - g) / (max - min) + 255 * 2 / 3;
                }
            }
            if (h >= (int)h + 0.5f)
            {
                h = (int)h + 1;
            }

            l = (max + min) / 2;
            if (max + min == 255)
            {
                l = 128;
            }
            if (l >= (int)l + 0.5f)
            {
                l = (int)l + 1;
            }

            if (l == 0 || max == min)
            {
                s = 0;
            }
            else
            {
                if (l <= 255 / 2)
                {
                    s = 255 * (max - min) / (max + min);
                }
                else
                {
                    s = 255 * (max - min) / (2 * 255 - (max + min));
                }
            }
            if (s >= (int)s + 0.5f)
            {
                s = (int)s + 1;
            }

            int hsl = (int)h + (int)s * 256 + (int)l * 256 * 256;
            return hsl;
        }

        private int Nrgb(int rgb)
        {
            string[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
            string type = Convert.ToString(arr[0]);
            int djz = int.Parse(arr[1]);

            int r0 = rgb % 256;
            int g0 = (rgb / 256) % 256;
            int b0 = (rgb / 256 / 256) % 256;
            int hsl = Rgb2Hsl(r0, g0, b0);
            int h0 = hsl % 256;
            int s0 = (hsl / 256) % 256;
            int l0 = (hsl / 256 / 256) % 256;
            int nrgb=0;
            if (type == "h")
            {
                 int h1 = h0 + djz;
                 if (h1 > 255)
                 {
                     h1 = 255;
                 }
                 else
                 {
                     if (h1 < 0)
                     {
                         h1 = 0;
                     }
                 }
                 label1.Text = "hsl:" + h1 + "," + s0 + "," + l0;
                 nrgb = Hsl2Rgb(h1, s0, l0);
            }
            if (type == "s")
            {
                 int s1 = s0 + djz;
                 if (s1 > 255)
                 {
                     s1 = 255;
                 }
                 else
                 {
                     if (s1 < 0)
                     {
                         s1 = 0;
                     }
                 }
                 label1.Text = "hsl:" + h0 + "," + s1 + "," + l0;
                 nrgb = Hsl2Rgb(h0, s1, l0);
            }
            if (type == "l")
            {
                 int l1 = l0 + djz;
                 if (l1 > 255)
                 {
                     l1 = 255;
                 }
                 else
                 {
                     if (l1 < 0)
                     {
                         l1 = 0;
                     }
                 }
                 label1.Text = "hsl:" + h0 + "," + s0 + "," + l1;
                 nrgb = Hsl2Rgb(h0, s0, l1);
            }
            if (type == "hs")
            {
                 int h1 = h0 + djz;
                 int s1 = s0 + djz;
                 if (h1 > 255)
                 {
                     h1 = 255;
                 }
                 else
                 {
                     if (h1 < 0)
                     {
                         h1 = 0;
                     }
                 }
                 if (s1 > 255)
                 {
                     s1 = 255;
                 }
                 else
                 {
                     if (s1 < 0)
                     {
                         s1 = 0;
                     }
                 }
                 label1.Text = "hsl:" + h1 + "," + s1 + "," + l0;
                 nrgb = Hsl2Rgb(h1, s1, l0);
            }
            if (type == "hl")
            {
                 int h1 = h0 + djz;
                 int l1 = l0 + djz;
                 if (h1 > 255)
                 {
                     h1 = 255;
                 }
                 else
                 {
                     if (h1 < 0)
                     {
                         h1 = 0;
                     }
                 }
                 if (l1 > 255)
                 {
                     l1 = 255;
                 }
                 else
                 {
                     if (l1 < 0)
                     {
                         l1 = 0;
                     }
                 }
                 label1.Text = "hsl:" + h1 + "," + s0 + "," + l1;
                 nrgb = Hsl2Rgb(h1, s0, l1);
            }
            if (type == "sl")
            {
                 int s1 = s0 + djz;
                 int l1 = l0 + djz;
                 if (s1 > 255)
                 {
                     s1 = 255;
                 }
                 else
                 {
                     if (s1 < 0)
                     {
                         s1 = 0;
                     }
                 }
                 if (l1 > 255)
                 {
                     l1 = 255;
                 }
                 else
                 {
                     if (l1 < 0)
                     {
                         l1 = 0;
                     }
                 }
                 label1.Text = "hsl:" + h0 + "," + s1 + "," + l1;
                 nrgb = Hsl2Rgb(h0, s1, l1);
            }
            if (type == "hsl")
            {
                 int h1 = h0 + djz;
                 int s1 = s0 + djz;
                 int l1 = l0 + djz;
                 if (h1 > 255)
                 {
                     h1 = 255;
                 }
                 else
                 {
                     if (h1 < 0)
                     {
                         h1 = 0;
                     }
                 }
                 if (s1 > 255)
                 {
                     s1 = 255;
                 }
                 else
                 {
                     if (s1 < 0)
                     {
                         s1 = 0;
                     }
                 }
                 if (l1 > 255)
                 {
                     l1 = 255;
                 }
                 else
                 {
                     if (l1 < 0)
                     {
                         l1 = 0;
                     }
                 }
                 label1.Text = "hsl:" + h1 + "," + s1 + "," + l1;
                 nrgb = Hsl2Rgb(h1, s1, l1);
            }
            if (type == "r")
            {
                 int r1 = r0 + djz;
                 if (r1 > 255)
                 {
                     r1 = 255;
                 }
                 else
                 {
                     if (r1 < 0)
                     {
                         r1 = 0;
                     }
                 }
                 label1.Text = "rgb:" + r1 + "," + g0 + "," + b0;
                 nrgb = r1 + g0 * 256 + b0 * 256 * 256;
            }
            if (type == "g")
            {
                 int g1 = g0 + djz;
                 if (g1 > 255)
                 {
                     g1 = 255;
                 }
                 else
                 {
                     if (g1 < 0)
                     {
                         g1 = 0;
                     }
                 }
                 label1.Text = "rgb:" + r0 + "," + g1 + "," + b0;
                 nrgb = r0 + g1 * 256 + b0 * 256 * 256;
            }
            if (type == "b")
            {
                 int b1 = b0 + djz;
                 if (b1 > 255)
                 {
                     b1 = 255;
                 }
                 else
                 {
                     if (b1 < 0)
                     {
                         b1 = 0;
                     }
                 }
                 label1.Text = "rgb:" + r0 + "," + g0 + "," + b1;
                 nrgb = r0 + g0 * 256 + b1 * 256 * 256;
            }
            if (type == "rg")
            {
                 int r1 = r0 + djz;
                 int g1 = g0 + djz;
                 if (r1 > 255)
                 {
                     r1 = 255;
                 }
                 else
                 {
                     if (r1 < 0)
                     {
                         r1 = 0;
                     }
                 }
                 if (g1 > 255)
                 {
                     g1 = 255;
                 }
                 else
                 {
                     if (g1 < 0)
                     {
                         g1 = 0;
                     }
                 }
                 label1.Text = "rgb:" + r1 + "," + g1 + "," + b0;
                 nrgb = r1 + g1 * 256 + b0 * 256 * 256;
            }
            if (type == "rb")
            {
                 int r1 = r0 + djz;
                 int b1 = b0 + djz;
                 if (r1 > 255)
                 {
                     r1 = 255;
                 }
                 else
                 {
                     if (r1 < 0)
                     {
                         r1 = 0;
                     }
                 }
                 if (b1 > 255)
                 {
                     b1 = 255;
                 }
                 else
                 {
                     if (b1 < 0)
                     {
                         b1 = 0;
                     }
                 }
                 label1.Text = "rgb:" + r1 + "," + g0 + "," + b1;
                 nrgb = r1 + g0 * 256 + b1 * 256 * 256;
            }
            if (type == "bg")
            {
                 int b1 = b0 + djz;
                 int g1 = g0 + djz;
                 if (g1 > 255)
                 {
                     g1 = 255;
                 }
                 else
                 {
                     if (g1 < 0)
                     {
                         g1 = 0;
                     }
                 }
                 if (b1 > 255)
                 {
                     b1 = 255;
                 }
                 else
                 {
                     if (b1 < 0)
                     {
                         b1 = 0;
                     }
                 }
                 label1.Text = "rgb:" + r0 + "," + g1 + "," + b1;
                 nrgb = r0 + g1 * 256 + b1 * 256 * 256;
            }
            if (type == "rgb")
            {
                 int r1 = r0 + djz;
                 int g1 = g0 + djz;
                 int b1 = b0 + djz;
                 if (r1 > 255)
                 {
                     r1 = 255;
                 }
                 else
                 {
                     if (r1 < 0)
                     {
                         r1 = 0;
                     }
                 }
                 if (g1 > 255)
                 {
                     g1 = 255;
                 }
                 else
                 {
                     if (g1 < 0)
                     {
                         g1 = 0;
                     }
                 }
                 if (b1 > 255)
                 {
                     b1 = 255;
                 }
                 else
                 {
                     if (b1 < 0)
                     {
                         b1 = 0;
                     }
                 }
                 label1.Text = "rgb:" + r1 + "," + g1 + "," + b1;
                 nrgb = r1 + g1 * 256 + b1 * 256 * 256;
            }
             return nrgb;
        }

        private int Nrgb2(int rgb,int a)
        {
            string[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
            string type = Convert.ToString(arr[0]);
            int djz = int.Parse(arr[1]) * a;

            int r0 = rgb % 256;
            int g0 = (rgb / 256) % 256;
            int b0 = (rgb / 256 / 256) % 256;
            int hsl = Rgb2Hsl(r0, g0, b0);
            int h0 = hsl % 256;
            int s0 = (hsl / 256) % 256;
            int l0 = (hsl / 256 / 256) % 256;
            int nrgb = 0;
            if (type == "h")
            {
                int h1 = h0 + djz;
                if (h1 > 255)
                {
                    h1 = 255;
                }
                else
                {
                    if (h1 < 0)
                    {
                        h1 = 0;
                    }
                }
                label1.Text = "hsl:" + h1 + "," + s0 + "," + l0;
                nrgb = Hsl2Rgb(h1, s0, l0);
            }
            if (type == "s")
            {
                int s1 = s0 + djz;
                if (s1 > 255)
                {
                    s1 = 255;
                }
                else
                {
                    if (s1 < 0)
                    {
                        s1 = 0;
                    }
                }
                label1.Text = "hsl:" + h0 + "," + s1 + "," + l0;
                nrgb = Hsl2Rgb(h0, s1, l0);
            }
            if (type == "l")
            {
                int l1 = l0 + djz;
                if (l1 > 255)
                {
                    l1 = 255;
                }
                else
                {
                    if (l1 < 0)
                    {
                        l1 = 0;
                    }
                }
                label1.Text = "hsl:" + h0 + "," + s0 + "," + l1;
                nrgb = Hsl2Rgb(h0, s0, l1);
            }
            if (type == "hs")
            {
                int h1 = h0 + djz;
                int s1 = s0 + djz;
                if (h1 > 255)
                {
                    h1 = 255;
                }
                else
                {
                    if (h1 < 0)
                    {
                        h1 = 0;
                    }
                }
                if (s1 > 255)
                {
                    s1 = 255;
                }
                else
                {
                    if (s1 < 0)
                    {
                        s1 = 0;
                    }
                }
                label1.Text = "hsl:" + h1 + "," + s1 + "," + l0;
                nrgb = Hsl2Rgb(h1, s1, l0);
            }
            if (type == "hl")
            {
                int h1 = h0 + djz;
                int l1 = l0 + djz;
                if (h1 > 255)
                {
                    h1 = 255;
                }
                else
                {
                    if (h1 < 0)
                    {
                        h1 = 0;
                    }
                }
                if (l1 > 255)
                {
                    l1 = 255;
                }
                else
                {
                    if (l1 < 0)
                    {
                        l1 = 0;
                    }
                }
                label1.Text = "hsl:" + h1 + "," + s0 + "," + l1;
                nrgb = Hsl2Rgb(h1, s0, l1);
            }
            if (type == "sl")
            {
                int s1 = s0 + djz;
                int l1 = l0 + djz;
                if (s1 > 255)
                {
                    s1 = 255;
                }
                else
                {
                    if (s1 < 0)
                    {
                        s1 = 0;
                    }
                }
                if (l1 > 255)
                {
                    l1 = 255;
                }
                else
                {
                    if (l1 < 0)
                    {
                        l1 = 0;
                    }
                }
                label1.Text = "hsl:" + h0 + "," + s1 + "," + l1;
                nrgb = Hsl2Rgb(h0, s1, l1);
            }
            if (type == "hsl")
            {
                int h1 = h0 + djz;
                int s1 = s0 + djz;
                int l1 = l0 + djz;
                if (h1 > 255)
                {
                    h1 = 255;
                }
                else
                {
                    if (h1 < 0)
                    {
                        h1 = 0;
                    }
                }
                if (s1 > 255)
                {
                    s1 = 255;
                }
                else
                {
                    if (s1 < 0)
                    {
                        s1 = 0;
                    }
                }
                if (l1 > 255)
                {
                    l1 = 255;
                }
                else
                {
                    if (l1 < 0)
                    {
                        l1 = 0;
                    }
                }
                label1.Text = "hsl:" + h1 + "," + s1 + "," + l1;
                nrgb = Hsl2Rgb(h1, s1, l1);
            }
            if (type == "r")
            {
                int r1 = r0 + djz;
                if (r1 > 255)
                {
                    r1 = 255;
                }
                else
                {
                    if (r1 < 0)
                    {
                        r1 = 0;
                    }
                }
                label1.Text = "rgb:" + r1 + "," + g0 + "," + b0;
                nrgb = r1 + g0 * 256 + b0 * 256 * 256;
            }
            if (type == "g")
            {
                int g1 = g0 + djz;
                if (g1 > 255)
                {
                    g1 = 255;
                }
                else
                {
                    if (g1 < 0)
                    {
                        g1 = 0;
                    }
                }
                label1.Text = "rgb:" + r0 + "," + g1 + "," + b0;
                nrgb = r0 + g1 * 256 + b0 * 256 * 256;
            }
            if (type == "b")
            {
                int b1 = b0 + djz;
                if (b1 > 255)
                {
                    b1 = 255;
                }
                else
                {
                    if (b1 < 0)
                    {
                        b1 = 0;
                    }
                }
                label1.Text = "rgb:" + r0 + "," + g0 + "," + b1;
                nrgb = r0 + g0 * 256 + b1 * 256 * 256;
            }
            if (type == "rg")
            {
                int r1 = r0 + djz;
                int g1 = g0 + djz;
                if (r1 > 255)
                {
                    r1 = 255;
                }
                else
                {
                    if (r1 < 0)
                    {
                        r1 = 0;
                    }
                }
                if (g1 > 255)
                {
                    g1 = 255;
                }
                else
                {
                    if (g1 < 0)
                    {
                        g1 = 0;
                    }
                }
                label1.Text = "rgb:" + r1 + "," + g1 + "," + b0;
                nrgb = r1 + g1 * 256 + b0 * 256 * 256;
            }
            if (type == "rb")
            {
                int r1 = r0 + djz;
                int b1 = b0 + djz;
                if (r1 > 255)
                {
                    r1 = 255;
                }
                else
                {
                    if (r1 < 0)
                    {
                        r1 = 0;
                    }
                }
                if (b1 > 255)
                {
                    b1 = 255;
                }
                else
                {
                    if (b1 < 0)
                    {
                        b1 = 0;
                    }
                }
                label1.Text = "rgb:" + r1 + "," + g0 + "," + b1;
                nrgb = r1 + g0 * 256 + b1 * 256 * 256;
            }
            if (type == "bg")
            {
                int b1 = b0 + djz;
                int g1 = g0 + djz;
                if (g1 > 255)
                {
                    g1 = 255;
                }
                else
                {
                    if (g1 < 0)
                    {
                        g1 = 0;
                    }
                }
                if (b1 > 255)
                {
                    b1 = 255;
                }
                else
                {
                    if (b1 < 0)
                    {
                        b1 = 0;
                    }
                }
                label1.Text = "rgb:" + r0 + "," + g1 + "," + b1;
                nrgb = r0 + g1 * 256 + b1 * 256 * 256;
            }
            if (type == "rgb")
            {
                int r1 = r0 + djz;
                int g1 = g0 + djz;
                int b1 = b0 + djz;
                if (r1 > 255)
                {
                    r1 = 255;
                }
                else
                {
                    if (r1 < 0)
                    {
                        r1 = 0;
                    }
                }
                if (g1 > 255)
                {
                    g1 = 255;
                }
                else
                {
                    if (g1 < 0)
                    {
                        g1 = 0;
                    }
                }
                if (b1 > 255)
                {
                    b1 = 255;
                }
                else
                {
                    if (b1 < 0)
                    {
                        b1 = 0;
                    }
                }
                label1.Text = "rgb:" + r1 + "," + g1 + "," + b1;
                nrgb = r1 + g1 * 256 + b1 * 256 * 256;
            }
            return nrgb;
        }

        //颜色单值统一转换代码
        private int Nrgb3(int rgb)
        {
            string[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
            string type = Convert.ToString(arr[0]);
            int nc = int.Parse(arr[1]);

            int r0 = rgb % 256;
            int g0 = (rgb / 256) % 256;
            int b0 = (rgb / 256 / 256) % 256;

            int h0 = 0; int s0 = 0; int l0 = 0;
            if (type == "h" || type == "s" || type == "l" || type == "hs" ||type =="hl" || type == "sl" || type == "hsl")
            {
                int hsl = Rgb2Hsl(r0, g0, b0);
                h0 = hsl % 256;
                s0 = (hsl / 256) % 256;
                l0 = (hsl / 256 / 256) % 256;
            }

            int nrgb = 0;
            if (type == "h")
            {
                int h1 = nc;
                if (h1 > 255)
                {
                    h1 = 255;
                }
                else if (h1 < 0)
                {
                    h1 = 0;
                }
                label1.Text = "hsl:" + h1 + "," + s0 + "," + l0;
                nrgb = Hsl2Rgb(h1, s0, l0);
            }
            if (type == "s")
            {
                int s1 = nc;
                if (s1 > 255)
                {
                    s1 = 255;
                }
                else if (s1 < 0)
                {
                    s1 = 0;
                }
                label1.Text = "hsl:" + h0 + "," + s1 + "," + l0;
                nrgb = Hsl2Rgb(h0, s1, l0);
            }
            if (type == "l")
            {
                int l1 = nc;
                if (l1 > 255)
                {
                    l1 = 255;
                }
                else if (l1 < 0)
                {
                    l1 = 0;
                }
                label1.Text = "hsl:" + h0 + "," + s0 + "," + l1;
                nrgb = Hsl2Rgb(h0, s0, l1);
            }
            if (type == "hs")
            {
                int h1 = nc;
                int s1 = nc;
                if (h1 > 255)
                {
                    h1 = 255;
                }
                else if (h1 < 0)
                {
                    h1 = 0;
                }
                if (s1 > 255)
                {
                    s1 = 255;
                }
                else if (s1 < 0)
                {
                    s1 = 0;
                }
                label1.Text = "hsl:" + h1 + "," + s1 + "," + l0;
                nrgb = Hsl2Rgb(h1, s1, l0);
            }
            if (type == "hl")
            {
                int h1 = nc;
                int l1 = nc;
                if (h1 > 255)
                {
                    h1 = 255;
                }
                else if (h1 < 0)
                {
                    h1 = 0;
                }
                if (l1 > 255)
                {
                    l1 = 255;
                }
                else if (l1 < 0)
                {
                    l1 = 0;
                }
                label1.Text = "hsl:" + h1 + "," + s0 + "," + l1;
                nrgb = Hsl2Rgb(h1, s0, l1);
            }
            if (type == "sl")
            {
                int s1 = nc;
                int l1 = nc;
                if (s1 > 255)
                {
                    s1 = 255;
                }
                else if (s1 < 0)
                {
                    s1 = 0;
                }
                if (l1 > 255)
                {
                    l1 = 255;
                }
                else if (l1 < 0)
                {
                    l1 = 0;
                }
                label1.Text = "hsl:" + h0 + "," + s1 + "," + l1;
                nrgb = Hsl2Rgb(h0, s1, l1);
            }
            if (type == "hsl")
            {
                int h1 = nc;
                int s1 = nc;
                int l1 = nc;
                if (h1 > 255)
                {
                    h1 = 255;
                }
                else if (h1 < 0)
                {
                    h1 = 0;
                }
                if (s1 > 255)
                {
                    s1 = 255;
                }
                else if (s1 < 0)
                {
                    s1 = 0;
                }
                if (l1 > 255)
                {
                    l1 = 255;
                }
                else if (l1 < 0)
                {
                    l1 = 0;
                }
                label1.Text = "hsl:" + h1 + "," + s1 + "," + l1;
                nrgb = Hsl2Rgb(h1, s1, l1);
            }
            if (type == "r")
            {
                int r1 = nc;
                if (r1 > 255)
                {
                    r1 = 255;
                }
                else if (r1 < 0)
                {
                    r1 = 0;
                }
                label1.Text = "rgb:" + r1 + "," + g0 + "," + b0;
                nrgb = r1 + g0 * 256 + b0 * 256 * 256;
            }
            if (type == "g")
            {
                int g1 = nc;
                if (g1 > 255)
                {
                    g1 = 255;
                }
                else if (g1 < 0)
                {
                    g1 = 0;
                }
                label1.Text = "rgb:" + r0 + "," + g1 + "," + b0;
                nrgb = r0 + g1 * 256 + b0 * 256 * 256;
            }
            if (type == "b")
            {
                int b1 = nc;
                if (b1 > 255)
                {
                    b1 = 255;
                }
                else if (b1 < 0)
                {
                    b1 = 0;
                }
                label1.Text = "rgb:" + r0 + "," + g0 + "," + b1;
                nrgb = r0 + g0 * 256 + b1 * 256 * 256;
            }
            if (type == "rg")
            {
                int r1 = nc;
                int g1 = nc;
                if (r1 > 255)
                {
                    r1 = 255;
                }
                else if (r1 < 0)
                {
                    r1 = 0;
                }
                if (g1 > 255)
                {
                    g1 = 255;
                }
                else if (g1 < 0)
                {
                    g1 = 0;
                }
                label1.Text = "rgb:" + r1 + "," + g1 + "," + b0;
                nrgb = r1 + g1 * 256 + b0 * 256 * 256;
            }
            if (type == "rb")
            {
                int r1 = nc;
                int b1 = nc;
                if (r1 > 255)
                {
                    r1 = 255;
                }
                else if (r1 < 0)
                {
                    r1 = 0;
                }
                if (b1 > 255)
                {
                    b1 = 255;
                }
                else if (b1 < 0)
                {
                    b1 = 0;
                }
                label1.Text = "rgb:" + r1 + "," + g0 + "," + b1;
                nrgb = r1 + g0 * 256 + b1 * 256 * 256;
            }
            if (type == "bg")
            {
                int b1 = nc;
                int g1 = nc;
                if (g1 > 255)
                {
                    g1 = 255;
                }
                else if (g1 < 0)
                {
                    g1 = 0;
                }
                if (b1 > 255)
                {
                    b1 = 255;
                }
                else if (b1 < 0)
                {
                    b1 = 0;
                }
                label1.Text = "rgb:" + r0 + "," + g1 + "," + b1;
                nrgb = r0 + g1 * 256 + b1 * 256 * 256;
            }
            if (type == "rgb")
            {
                int r1 = nc;
                int g1 = nc;
                int b1 = nc;
                if (r1 > 255)
                {
                    r1 = 255;
                }
                else if (r1 < 0)
                {
                    r1 = 0;
                }
                if (g1 > 255)
                {
                    g1 = 255;
                }
                else if (g1 < 0)
                {
                    g1 = 0;
                }
                if (b1 > 255)
                {
                    b1 = 255;
                }
                else if (b1 < 0)
                {
                    b1 = 0;
                }
                label1.Text = "rgb:" + r1 + "," + g1 + "," + b1;
                nrgb = r1 + g1 * 256 + b1 * 256 * 256;
            }
            return nrgb;
        }

        public bool bolnum(string temp)
        {
            for (int i = 0; i < temp.Length; i++)
            {
                byte tempByte = Convert.ToByte(temp[i]);
                if (tempByte < 48 || tempByte > 57)
                {
                    return false;
                }
            }
            return true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OK_Fast.ActiveForm.Close();
        }

        private PowerPoint.Application app = Globals.ThisAddIn.Application;

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            
            if (comboBox1.SelectedItem as string == "数值上色(RGB)")
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    label1.Text = "先选矢量形状";
                }
                else
                {
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    if (sel.HasChildShapeRange)
                    {
                        range = sel.ChildShapeRange;
                    }
                    if (range[1].Fill.Type == Office.MsoFillType.msoFillSolid || range[1].Fill.Type == Office.MsoFillType.msoFillGradient)
                    {
                        int rgb0 = range[1].Fill.ForeColor.RGB;
                        int r0 = rgb0 % 256;
                        int g0 = (rgb0 / 256) % 256;
                        int b0 = (rgb0 / 256 / 256) % 256;
                        label1.Text = r0 + "," + g0 + "," + b0;
                        String[] arr = textBox1.Text.Trim().Split(char.Parse(","), char.Parse(" "), char.Parse("，")).ToArray();
                        int r = int.Parse(arr[0]);
                        int g = int.Parse(arr[1]);
                        int b = int.Parse(arr[2]);
                        int count = range.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            range[i].Fill.ForeColor.RGB = r + g * 256 + b * 256 * 256;
                        }
                        label1.Text = r + "," + g + "," + b;
                    }
                    else
                    {
                        label1.Text = "需选纯色形状";
                    }
                    
                }
            }

            if (comboBox1.SelectedItem as string == "数值上色(HSL)")
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    label1.Text = "先选矢量形状";
                }
                else
                {
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    if (sel.HasChildShapeRange)
                    {
                        range = sel.ChildShapeRange;
                    }
                    if (range[1].Fill.Type == Office.MsoFillType.msoFillSolid || range[1].Fill.Type == Office.MsoFillType.msoFillGradient)
                    {
                        int rgb0 = range[1].Fill.ForeColor.RGB;
                        int r0 = rgb0 % 256;
                        int g0 = (rgb0 / 256) % 256;
                        int b0 = (rgb0 / 256 / 256) % 256;
                        int hsl = Rgb2Hsl(r0, g0, b0);
                        int h0 = hsl % 256;
                        int s0 = (hsl / 256) % 256;
                        int l0 = (hsl / 256 / 256) % 256;
                        label1.Text = h0 + "," + s0 + "," + l0;
                        String[] arr = textBox1.Text.Trim().Split(char.Parse(","), char.Parse(" "), char.Parse("，")).ToArray();
                        int h = int.Parse(arr[0]);
                        int s = int.Parse(arr[1]);
                        int l = int.Parse(arr[2]);
                        int rgb = Hsl2Rgb(h, s, l);
                        int count = range.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            range[i].Fill.ForeColor.RGB = rgb;
                        }
                        label1.Text = h + "," + s + "," + l;
                    }
                    else
                    {
                        label1.Text = "需选纯色形状";
                    }
                }
            }

            if (comboBox1.SelectedItem as string == "数值上色(16进制)")
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    label1.Text = "先选矢量形状";
                }
                else
                {
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    if (sel.HasChildShapeRange)
                    {
                        range = sel.ChildShapeRange;
                    }
                    if (range[1].Fill.Type == Office.MsoFillType.msoFillSolid || range[1].Fill.Type == Office.MsoFillType.msoFillGradient)
                    {
                        int rgb0 = range[1].Fill.ForeColor.RGB;
                        int r0 = rgb0 % 256;
                        int g0 = (rgb0 / 256) % 256;
                        int b0 = (rgb0 / 256 / 256) % 256;
                        label1.Text = r0 + "," + g0 + "," + b0;

                        string hex = textBox1.Text.Trim();
                        Color rgb = ColorTranslator.FromHtml(hex);
                        int count = range.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            range[i].Fill.ForeColor.RGB = rgb.R + rgb.G * 256 + rgb.B * 256 * 256;
                        }
                        label1.Text = rgb.R + "," + rgb.G + "," + rgb.B;
                    }
                    else
                    {
                        label1.Text = "需选纯色形状";
                    }

                }
            }

            if (comboBox1.SelectedItem as string == "旋转递进")
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    label1.Text = "起始值,终止值";
                }
                else
                {
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    if (sel.HasChildShapeRange)
                    {
                        range = sel.ChildShapeRange;
                    }
                    int count = range.Count;
                    if (count == 1)
                    {
                        label1.Text = range[1].Rotation + "°";
                        range[1].Rotation = float.Parse(textBox1.Text);
                        label1.Text = Convert.ToString(range[1].Rotation) + "°";
                    }
                    else
                    {
                        if (count >= 2)
                        {
                            float r0 = range[1].Rotation;
                            float rc0 = range[count].Rotation;
                            label1.Text = "起始" + r0 + "°,终止" + rc0 + "°";
                            String[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse("，"), char.Parse(",")).ToArray();
                            float r1 = float.Parse(arr[0]);
                            float rc1 = float.Parse(arr[1]);
                            float n = (rc1 - r1) / (count - 1);
                            for (int i = 1; i <= count; i++)
                            {
                                range[i].Rotation = r1 + n * (i - 1);
                            }
                            label1.Text = "起始" + range[1].Rotation + "°,终止" + range[count].Rotation + "°";
                        }
                    } 
                }
            }

            if (comboBox1.SelectedItem as string == "本色渐变(H)")
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    label1.Text = "色调差";
                }
                else
                {
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    if (sel.HasChildShapeRange)
                    {
                        range = sel.ChildShapeRange;
                    }
                    int count = range.Count;

                    if (textBox1.Text.Trim() != "")
                    {
                        int nc = int.Parse(textBox1.Text.Trim());
                        label1.Text = "色调差是：" + nc;
                        for (int i = 1; i <= count; i++)
                        {
                            if (range[i].Type == Office.MsoShapeType.msoGroup)
                            {
                                for (int j = 1; j <= range[i].GroupItems.Count; j++)
                                {
                                    PowerPoint.Shape shape = range[i].GroupItems[j];
                                    if (shape.Fill.Type == Office.MsoFillType.msoFillSolid)
                                    {
                                        int rgb = shape.Fill.ForeColor.RGB;
                                        int r = rgb % 256;
                                        int g = (rgb / 256) % 256;
                                        int b = (rgb / 256 / 256) % 256;
                                        int hsl = Rgb2Hsl(r, g, b);
                                        int h = hsl % 256;
                                        int s = (hsl / 256) % 256;
                                        int l = (hsl / 256 / 256) % 256;
                                        int nh = h + nc;
                                        if (nh > 255)
                                        {
                                            nh = nh - 256;
                                        }
                                        else if (nh < 0)
                                        {
                                            nh = 256 - nh;
                                        }
                                        int nrgb = Hsl2Rgb(nh, s, l);
                                        shape.Fill.OneColorGradient(Office.MsoGradientStyle.msoGradientDiagonalUp, 1, 1);
                                        shape.Fill.GradientStops[1].Color.RGB = rgb;
                                        shape.Fill.GradientStops[2].Color.RGB = nrgb;
                                        shape.Fill.GradientAngle = 0;
                                    }
                                    else if (shape.Fill.Type == Office.MsoFillType.msoFillGradient)
                                    {
                                        int rgb = shape.Fill.GradientStops[1].Color.RGB;
                                        int r = rgb % 256;
                                        int g = (rgb / 256) % 256;
                                        int b = (rgb / 256 / 256) % 256;
                                        int hsl = Rgb2Hsl(r, g, b);
                                        int h = hsl % 256;
                                        int s = (hsl / 256) % 256;
                                        int l = (hsl / 256 / 256) % 256;
                                        int nh = h + nc;
                                        if (nh > 255)
                                        {
                                            nh = nh - 256;
                                        }
                                        else if (nh < 0)
                                        {
                                            nh = 256 - nh;
                                        }
                                        int nrgb = Hsl2Rgb(nh, s, l);
                                        shape.Fill.GradientStops[2].Color.RGB = nrgb;
                                    }
                                    else
                                    {
                                        label1.Text = "所选非渐变";
                                    }
                                }
                            }
                            else
	                        {
                                PowerPoint.Shape shape = range[i];
                                if (shape.Fill.Type == Office.MsoFillType.msoFillSolid)
                                {
                                    int rgb = shape.Fill.ForeColor.RGB;
                                    int r = rgb % 256;
                                    int g = (rgb / 256) % 256;
                                    int b = (rgb / 256 / 256) % 256;
                                    int hsl = Rgb2Hsl(r, g, b);
                                    int h = hsl % 256;
                                    int s = (hsl / 256) % 256;
                                    int l = (hsl / 256 / 256) % 256;
                                    int nh = h + nc;
                                    if (nh > 255)
                                    {
                                        nh = nh - 256;
                                    }
                                    else if (nh < 0)
                                    {
                                        nh = 256 - nh;
                                    }
                                    int nrgb = Hsl2Rgb(nh, s, l);
                                    shape.Fill.OneColorGradient(Office.MsoGradientStyle.msoGradientDiagonalUp, 1, 1);
                                    shape.Fill.GradientStops[1].Color.RGB = rgb;
                                    shape.Fill.GradientStops[2].Color.RGB = nrgb;
                                    shape.Fill.GradientAngle = 0;
                                }
                                else if (shape.Fill.Type == Office.MsoFillType.msoFillGradient)
                                {
                                    int rgb = shape.Fill.GradientStops[1].Color.RGB;
                                    int r = rgb % 256;
                                    int g = (rgb / 256) % 256;
                                    int b = (rgb / 256 / 256) % 256;
                                    int hsl = Rgb2Hsl(r, g, b);
                                    int h = hsl % 256;
                                    int s = (hsl / 256) % 256;
                                    int l = (hsl / 256 / 256) % 256;
                                    int nh = h + nc;
                                    if (nh > 255)
                                    {
                                        nh = nh - 256;
                                    }
                                    else if (nh < 0)
                                    {
                                        nh = 256 - nh;
                                    }
                                    int nrgb = Hsl2Rgb(nh, s, l);
                                    shape.Fill.GradientStops[2].Color.RGB = nrgb;
                                }
                                else
                                {
                                    label1.Text = "所选非渐变";
                                }
                            }
                        }
                    }
                }
            }

            if (comboBox1.SelectedItem as string == "本色渐变(S)")
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    label1.Text = "饱和差";
                }
                else
                {
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    if (sel.HasChildShapeRange)
                    {
                        range = sel.ChildShapeRange;
                    }
                    int count = range.Count;
                    if (textBox1.Text.Trim() != "")
                    {
                        int nc = int.Parse(textBox1.Text.Trim());
                        label1.Text = "饱和差是：" + nc;
                        for (int i = 1; i <= count; i++)
                        {
                            if (range[i].Type == Office.MsoShapeType.msoGroup)
                            {
                                for (int j = 1; j <= range[i].GroupItems.Count; j++)
                                {
                                    PowerPoint.Shape shape = range[i].GroupItems[j];
                                    if (shape.Fill.Type == Office.MsoFillType.msoFillSolid)
                                    {
                                        int rgb = shape.Fill.ForeColor.RGB;
                                        int r = rgb % 256;
                                        int g = (rgb / 256) % 256;
                                        int b = (rgb / 256 / 256) % 256;
                                        int hsl = Rgb2Hsl(r, g, b);
                                        int h = hsl % 256;
                                        int s = (hsl / 256) % 256;
                                        int l = (hsl / 256 / 256) % 256;
                                        int ns = 0;
                                        if (s <= 255 - nc && s + nc >= 0)
                                        {
                                            ns = s + nc;
                                            int nrgb = Hsl2Rgb(h, ns, l);
                                            shape.Fill.OneColorGradient(Office.MsoGradientStyle.msoGradientDiagonalUp, 1, 1);
                                            shape.Fill.GradientStops[1].Color.RGB = rgb;
                                            shape.Fill.GradientStops[2].Color.RGB = nrgb;
                                            shape.Fill.GradientAngle = 90;
                                        }
                                        else
                                        {
                                            label1.Text = "已经最艳/灰";
                                        }
                                    }
                                    else if (shape.Fill.Type == Office.MsoFillType.msoFillGradient)
                                    {
                                        int rgb = shape.Fill.GradientStops[1].Color.RGB;
                                        int r = rgb % 256;
                                        int g = (rgb / 256) % 256;
                                        int b = (rgb / 256 / 256) % 256;
                                        int hsl = Rgb2Hsl(r, g, b);
                                        int h = hsl % 256;
                                        int s = (hsl / 256) % 256;
                                        int l = (hsl / 256 / 256) % 256;
                                        int ns = 0;
                                        if (s <= 255 - nc && s + nc >= 0)
                                        {
                                            ns = s + nc;
                                            int nrgb = Hsl2Rgb(h, ns, l);
                                            shape.Fill.GradientStops[2].Color.RGB = nrgb;
                                        }
                                        else
                                        {
                                            label1.Text = "已经最艳/灰";
                                        }
                                    }
                                    else
                                    {
                                        label1.Text = "所选非渐变";
                                    }
                                }
                            }
                            else
                            {
                                PowerPoint.Shape shape = range[i];
                                if (shape.Fill.Type == Office.MsoFillType.msoFillSolid)
                                {
                                    int rgb = shape.Fill.ForeColor.RGB;
                                    int r = rgb % 256;
                                    int g = (rgb / 256) % 256;
                                    int b = (rgb / 256 / 256) % 256;
                                    int hsl = Rgb2Hsl(r, g, b);
                                    int h = hsl % 256;
                                    int s = (hsl / 256) % 256;
                                    int l = (hsl / 256 / 256) % 256;
                                    int ns = 0;
                                    if (s <= 255 - nc && s + nc >= 0)
                                    {
                                        ns = s + nc;
                                        int nrgb = Hsl2Rgb(h, ns, l);
                                        shape.Fill.OneColorGradient(Office.MsoGradientStyle.msoGradientDiagonalUp, 1, 1);
                                        shape.Fill.GradientStops[1].Color.RGB = rgb;
                                        shape.Fill.GradientStops[2].Color.RGB = nrgb;
                                        shape.Fill.GradientAngle = 90;
                                    }
                                    else
                                    {
                                        label1.Text = "已经最艳/灰";
                                    }
                                }
                                else if (shape.Fill.Type == Office.MsoFillType.msoFillGradient)
                                {
                                    int rgb = shape.Fill.GradientStops[1].Color.RGB;
                                    int r = rgb % 256;
                                    int g = (rgb / 256) % 256;
                                    int b = (rgb / 256 / 256) % 256;
                                    int hsl = Rgb2Hsl(r, g, b);
                                    int h = hsl % 256;
                                    int s = (hsl / 256) % 256;
                                    int l = (hsl / 256 / 256) % 256;
                                    int ns = 0;
                                    if (s <= 255 - nc && s + nc >= 0)
                                    {
                                        ns = s + nc;
                                        int nrgb = Hsl2Rgb(h, ns, l);
                                        shape.Fill.GradientStops[2].Color.RGB = nrgb;
                                    }
                                    else
                                    {
                                        label1.Text = "已经最艳/灰";
                                    }
                                }
                                else
                                {
                                    label1.Text = "所选非渐变";
                                }
                            }
                        }
                    }
                }
            }

            if (comboBox1.SelectedItem as string == "本色渐变(L)")
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    label1.Text = "亮度差";
                }
                else
                {
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    if (sel.HasChildShapeRange)
                    {
                        range = sel.ChildShapeRange;
                    }
                    int count = range.Count;
                    if (textBox1.Text.Trim() != "")
                    {
                        int nc = int.Parse(textBox1.Text.Trim());
                        label1.Text = "亮度差是：" + nc;
                        for (int i = 1; i <= count; i++)
                        {
                            if (range[i].Type == Office.MsoShapeType.msoGroup)
                            {
                                for (int j = 1; j <= range[i].GroupItems.Count; j++)
                                {
                                    PowerPoint.Shape shape = range[i].GroupItems[j];
                                    if (shape.Fill.Type == Office.MsoFillType.msoFillSolid)
                                    {
                                        int rgb = shape.Fill.ForeColor.RGB;
                                        int r = rgb % 256;
                                        int g = (rgb / 256) % 256;
                                        int b = (rgb / 256 / 256) % 256;
                                        int hsl = Rgb2Hsl(r, g, b);
                                        int h = hsl % 256;
                                        int s = (hsl / 256) % 256;
                                        int l = (hsl / 256 / 256) % 256;
                                        int nl = 0;
                                        if (l <= 255 - nc && l + nc >= 0)
                                        {
                                            nl = l + nc;
                                            int nrgb = Hsl2Rgb(h, s, nl);
                                            shape.Fill.OneColorGradient(Office.MsoGradientStyle.msoGradientDiagonalUp, 1, 1);
                                            shape.Fill.GradientStops[1].Color.RGB = rgb;
                                            shape.Fill.GradientStops[2].Color.RGB = nrgb;
                                            shape.Fill.GradientAngle = 90;
                                        }
                                        else
                                        {
                                            label1.Text = "已经最亮/暗";
                                        }
                                    }
                                    else if (shape.Fill.Type == Office.MsoFillType.msoFillGradient)
                                    {
                                        int rgb = shape.Fill.GradientStops[1].Color.RGB;
                                        int r = rgb % 256;
                                        int g = (rgb / 256) % 256;
                                        int b = (rgb / 256 / 256) % 256;
                                        int hsl = Rgb2Hsl(r, g, b);
                                        int h = hsl % 256;
                                        int s = (hsl / 256) % 256;
                                        int l = (hsl / 256 / 256) % 256;
                                        int nl = 0;
                                        if (l <= 255 - nc && l + nc >= 0)
                                        {
                                            nl = l + nc;
                                            int nrgb = Hsl2Rgb(h, s, nl);
                                            shape.Fill.GradientStops[2].Color.RGB = nrgb;
                                        }
                                        else
                                        {
                                            label1.Text = "已经最亮/暗";
                                        }
                                    }
                                    else
                                    {
                                        label1.Text = "所选非渐变";
                                    }
                                } 
                            }
                            else
                            {
                                PowerPoint.Shape shape = range[i];
                                if (shape.Fill.Type == Office.MsoFillType.msoFillSolid)
                                {
                                    int rgb = shape.Fill.ForeColor.RGB;
                                    int r = rgb % 256;
                                    int g = (rgb / 256) % 256;
                                    int b = (rgb / 256 / 256) % 256;
                                    int hsl = Rgb2Hsl(r, g, b);
                                    int h = hsl % 256;
                                    int s = (hsl / 256) % 256;
                                    int l = (hsl / 256 / 256) % 256;
                                    int nl = 0;
                                    if (l <= 255 - nc && l + nc >= 0)
                                    {
                                        nl = l + nc;
                                        int nrgb = Hsl2Rgb(h, s, nl);
                                        shape.Fill.OneColorGradient(Office.MsoGradientStyle.msoGradientDiagonalUp, 1, 1);
                                        shape.Fill.GradientStops[1].Color.RGB = rgb;
                                        shape.Fill.GradientStops[2].Color.RGB = nrgb;
                                        shape.Fill.GradientAngle = 90;
                                    }
                                    else
                                    {
                                        label1.Text = "已经最亮/暗";
                                    }
                                }
                                else if (shape.Fill.Type == Office.MsoFillType.msoFillGradient)
                                {
                                    int rgb = shape.Fill.GradientStops[1].Color.RGB;
                                    int r = rgb % 256;
                                    int g = (rgb / 256) % 256;
                                    int b = (rgb / 256 / 256) % 256;
                                    int hsl = Rgb2Hsl(r, g, b);
                                    int h = hsl % 256;
                                    int s = (hsl / 256) % 256;
                                    int l = (hsl / 256 / 256) % 256;
                                    int nl = 0;
                                    if (l <= 255 - nc && l + nc >= 0)
                                    {
                                        nl = l + nc;
                                        int nrgb = Hsl2Rgb(h, s, nl);
                                        shape.Fill.GradientStops[2].Color.RGB = nrgb;
                                    }
                                    else
                                    {
                                        label1.Text = "已经最亮/暗";
                                    }
                                }
                                else
                                {
                                    label1.Text = "所选非渐变";
                                }
                            }
                        }
                    }  
                }
            }

            if (comboBox1.SelectedItem as string == "高宽递进")
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                     label1.Text = "高度差,宽度差";
                }
                else
                {
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    if (sel.HasChildShapeRange)
                    {
                        range = sel.ChildShapeRange;
                    }
                    int count = range.Count;
                    string[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
                    float hn0 = float.Parse(arr[0]);
                    float wn0 = float.Parse(arr[1]);
                    double hn1 = hn0 * 72 / 2.54;
                    double wn1 = wn0 * 72 / 2.54;

                    if (count == 1)
                    {
                        float h1 = range[1].Height;
                        float w1 = range[1].Width;
                        float l1 = range[1].Left;
                        float t1 = range[1].Top;
                        range[1].Height = (float)hn1;
                        range[1].Width = (float)wn1;
                        range[1].Left = l1 + w1 / 2 - range[1].Width / 2;
                        range[1].Top = t1 + h1 / 2 - range[1].Height / 2;
                        label1.Text = "高" + hn0 + ",宽" + wn0;
                    }
                    else
                    {
                        label1.Text = "高差" + Math.Round(hn0, 2) + "," + "宽差" + Math.Round(wn0, 2);
                        for (int i = 2; i <= count; i++)
                        {
                            PowerPoint.Shape shape = range[i];
                            float ol = range[i].Left;
                            float ow = range[i].Width;
                            float ot = range[i].Top;
                            float oh = range[i].Height;
                            range[i].Height = range[1].Height + (float)hn1 * (i - 1);
                            range[i].Top = ot + oh / 2 - range[i].Height / 2;
                            range[i].Width = range[1].Width + (float)wn1 * (i - 1);
                            range[i].Left = ol + ow / 2 - range[i].Width / 2;
                        }
                    }
                }
            }

            if (comboBox1.SelectedItem as string == "填充透明递进")
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    label1.Text = "起始值,终止值";
                }
                else
                {
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    if (sel.HasChildShapeRange)
                    {
                        range = sel.ChildShapeRange;
                    }
                    int count = range.Count;
                    string[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
                    if (arr.Count() == 1)
                    {
                        float tn0 = float.Parse(arr[0]) / 100;
                        for (int i = 1; i <= count; i++)
                        {
                            PowerPoint.Shape shape = range[i];
                            if (shape.Type == Office.MsoShapeType.msoGroup)
                            {
                                foreach (PowerPoint.Shape item in shape.GroupItems)
                                {
                                    if (item.Fill.Type == Office.MsoFillType.msoFillSolid || item.Fill.Type == Office.MsoFillType.msoFillPicture)
                                    {
                                        item.Fill.Transparency = tn0;
                                    }
                                    else if (item.Fill.Type == Office.MsoFillType.msoFillGradient)
                                    {
                                        for (int j = 1; j <= item.Fill.GradientStops.Count; j++)
                                        {
                                            item.Fill.GradientStops[j].Transparency = tn0;
                                        }
                                    }
                                }
                                label1.Text = "透明度是" + tn0 * 100 + "%";
                            }
                            else
                            {
                                if (shape.Fill.Type == Office.MsoFillType.msoFillSolid || shape.Fill.Type == Office.MsoFillType.msoFillPicture)
                                {
                                    shape.Fill.Transparency = tn0;
                                }
                                else if (shape.Fill.Type == Office.MsoFillType.msoFillGradient)
                                {
                                    for (int j = 1; j <= shape.Fill.GradientStops.Count; j++)
                                    {
                                        shape.Fill.GradientStops[j].Transparency = tn0;
                                    }
                                }
                                label1.Text = "透明度是" + tn0 * 100 + "%";
                            }
                        }
                    }
                    else
                    {
                        float tn0 = float.Parse(arr[0]) / 100;
                        float tn1 = float.Parse(arr[1]) / 100;
                        if (count == 1 && range[1].Type == Office.MsoShapeType.msoGroup)
                        {
                            float tn = (tn1 - tn0) / (range[1].GroupItems.Count - 1);
                            for (int i = 1; i <= range[1].GroupItems.Count; i++)
                            {
                                PowerPoint.Shape shape = range[1].GroupItems[i];
                                if (shape.Fill.Type == Office.MsoFillType.msoFillSolid || shape.Fill.Type == Office.MsoFillType.msoFillPicture)
                                {
                                    shape.Fill.Transparency = tn0 + tn * (i - 1);
                                }
                                else if (shape.Fill.Type == Office.MsoFillType.msoFillGradient)
                                {
                                    for (int j = 1; j <= shape.Fill.GradientStops.Count; j++)
                                    {
                                        shape.Fill.GradientStops[j].Transparency = tn0 + tn * (i - 1);
                                    }
                                }
                                label1.Text = "从" + tn0 * 100 + "%到" + tn1 * 100 + "%";
                            }
                        }
                        else
                        {
                            float tn = (tn1 - tn0) / (count - 1);
                            for (int i = 1; i <= count; i++)
                            {
                                PowerPoint.Shape shape = range[i];
                                if (shape.Type == Office.MsoShapeType.msoGroup)
                                {
                                    foreach (PowerPoint.Shape item in shape.GroupItems)
                                    {
                                        if (item.Fill.Type == Office.MsoFillType.msoFillSolid || item.Fill.Type == Office.MsoFillType.msoFillPicture)
                                        {
                                            item.Fill.Transparency = tn0 + tn * (i - 1);
                                        }
                                        else if (item.Fill.Type == Office.MsoFillType.msoFillGradient)
                                        {
                                            for (int j = 1; j <= item.Fill.GradientStops.Count; j++)
                                            {
                                                item.Fill.GradientStops[j].Transparency = tn0 + tn * (i - 1);
                                            }
                                        }
                                    }
                                    label1.Text = "从" + tn0 * 100 + "%到" + tn1 * 100 + "%";
                                }
                                else
                                {
                                    if (shape.Fill.Type == Office.MsoFillType.msoFillSolid || shape.Fill.Type == Office.MsoFillType.msoFillPicture)
                                    {
                                        shape.Fill.Transparency = tn0 + tn * (i - 1);
                                    }
                                    else if (shape.Fill.Type == Office.MsoFillType.msoFillGradient)
                                    {
                                        for (int j = 1; j <= shape.Fill.GradientStops.Count; j++)
                                        {
                                            shape.Fill.GradientStops[j].Transparency = tn0 + tn * (i - 1);
                                        }
                                    }
                                    label1.Text = "从" + tn0 * 100 + "%到" + tn1 * 100 + "%";
                                }
                            }
                        }         
                    }  
                }  
            }

            if (comboBox1.SelectedItem as string == "线条透明递进")
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    label1.Text = "起始值,终止值";
                }
                else
                {
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    if (sel.HasChildShapeRange)
                    {
                        range = sel.ChildShapeRange;
                    }
                    int count = range.Count;
                    string[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
                    if (arr.Count() == 1)
                    {
                        float tn0 = float.Parse(arr[0]) / 100;
                        for (int i = 1; i <= count; i++)
                        {
                            PowerPoint.Shape shape = range[i];
                            if (shape.Type == Office.MsoShapeType.msoGroup)
                            {
                                foreach (PowerPoint.Shape item in shape.GroupItems)
                                {
                                    if (item.Line.Visible == Office.MsoTriState.msoFalse)
                                    {
                                        item.Line.Visible = Office.MsoTriState.msoTrue;
                                    }
                                    item.Line.Transparency = tn0;
                                }
                                label1.Text = "透明度是" + tn0 * 100 + "%";
                            }
                            else
                            {
                                if (shape.Line.Visible == Office.MsoTriState.msoFalse)
                                {
                                    shape.Line.Visible = Office.MsoTriState.msoTrue;
                                }
                                shape.Line.Transparency = tn0;
                                label1.Text = "透明度是" + tn0 * 100 + "%";
                            }
                        }
                    }
                    else
                    {
                        float tn0 = float.Parse(arr[0]) / 100;
                        float tn1 = float.Parse(arr[1]) / 100;
                        if (count == 1 && range[1].Type == Office.MsoShapeType.msoGroup)
                        {
                            float tn = (tn1 - tn0) / (range[1].GroupItems.Count - 1);
                            for (int i = 1; i <= range[1].GroupItems.Count; i++)
                            {
                                PowerPoint.Shape shape = range[1].GroupItems[i];
                                if (shape.Line.Visible == Office.MsoTriState.msoFalse)
                                {
                                    shape.Line.Visible = Office.MsoTriState.msoTrue;
                                }
                                shape.Line.Transparency = tn0 + tn * (i - 1);
                            }
                            label1.Text = "从" + tn0 * 100 + "%到" + tn1 * 100 + "%";
                        }
                        else
                        {
                            float tn = (tn1 - tn0) / (count - 1);
                            for (int i = 1; i <= count; i++)
                            {
                                PowerPoint.Shape shape = range[i];
                                if (shape.Type == Office.MsoShapeType.msoGroup)
                                {
                                    foreach (PowerPoint.Shape item in shape.GroupItems)
                                    {
                                        if (item.Line.Visible == Office.MsoTriState.msoFalse)
                                        {
                                            item.Line.Visible = Office.MsoTriState.msoTrue;
                                        }
                                        item.Line.Transparency = tn0 + tn * (i - 1);
                                    }
                                    label1.Text = "从" + tn0 * 100 + "%到" + tn1 * 100 + "%";
                                }
                                else
                                {
                                    if (shape.Line.Visible == Office.MsoTriState.msoFalse)
                                    {
                                        shape.Line.Visible = Office.MsoTriState.msoTrue;
                                    }
                                    shape.Line.Transparency = tn0 + tn * (i - 1);
                                    label1.Text = "从" + tn0 * 100 + "%到" + tn1 * 100 + "%";
                                }
                            }
                        }
                    }
                }
            }

            if (comboBox1.SelectedItem as string == "文字透明递进")
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    label1.Text = "起始值,终止值";
                }
                else
                {
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    if (sel.HasChildShapeRange)
                    {
                        range = sel.ChildShapeRange;
                    }
                    int count = range.Count;
                    string[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
                    if (arr.Count() == 1)
                    {
                        float tn0 = float.Parse(arr[0]) / 100;
                        if (count == 1)
                        {
                            if (range[1].Type == Office.MsoShapeType.msoGroup)
                            {
                                foreach (PowerPoint.Shape item in range[1].GroupItems)
                                {
                                    if (item.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                                    {
                                        item.TextFrame2.TextRange.Characters.Font.Fill.Transparency = tn0;
                                    }
                                }
                            }
                            else
                            {
                                PowerPoint.TextFrame2 tf2 = range[1].TextFrame2;
                                int tcount = tf2.TextRange.Characters.Count;
                                for (int j = 1; j <= tcount; j++)
                                {
                                    tf2.TextRange.Characters.Font.Fill.Transparency = tn0;
                                }
                            }
                        }
                        else
                        {
                            for (int i = 1; i <= count; i++)
                            {
                                PowerPoint.Shape shape = range[i];
                                if (shape.Type == Office.MsoShapeType.msoGroup)
                                {
                                    foreach (PowerPoint.Shape item in shape.GroupItems)
                                    {
                                        if (item.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                                        {
                                            item.TextFrame2.TextRange.Characters.Font.Fill.Transparency = tn0;
                                        }
                                    }
                                }
                                else
                                {
                                    if (shape.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                                    {
                                        shape.TextFrame2.TextRange.Characters.Font.Fill.Transparency = tn0;
                                    }
                                }
                            }
                        }
                        label1.Text = "透明度是" + tn0 * 100 + "%";
                    }
                    else
                    {
                        float tn0 = float.Parse(arr[0]) / 100;
                        float tn1 = float.Parse(arr[1]) / 100;
                        if (count == 1)
                        {
                            if (range[1].Type == Office.MsoShapeType.msoGroup)
                            {
                                float tn = (tn1 - tn0) / (range[1].GroupItems.Count - 1);
                                for (int i = 1; i <= range[1].GroupItems.Count; i++)
                                {
                                    PowerPoint.Shape shape = range[1].GroupItems[i];
                                    if (shape.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                                    {
                                        shape.TextFrame2.TextRange.Font.Fill.Transparency = tn0 + tn * (i - 1);
                                    }
                                }
                            }
                            else
                            {
                                PowerPoint.TextFrame2 tf2 = range[1].TextFrame2;
                                int tcount = tf2.TextRange.Characters.Count;
                                float tn = (tn1 - tn0) / (tcount - 1);
                                for (int j = 1; j <= tcount; j++)
                                {
                                    tf2.TextRange.Characters[j].Font.Fill.Transparency = tn0 + tn * (j - 1);
                                }
                            }
                            label1.Text = "从" + tn0 * 100 + "%到" + tn1 * 100 + "%";
                        }
                        else
                        {
                            float tn = (tn1 - tn0) / (count - 1);
                            for (int i = 1; i <= count; i++)
                            {
                                PowerPoint.Shape shape = range[i];
                                if (shape.Type == Office.MsoShapeType.msoGroup)
                                {
                                    foreach (PowerPoint.Shape item in shape.GroupItems)
                                    {
                                        if (item.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                                        {
                                            item.TextFrame2.TextRange.Font.Fill.Transparency = tn0 + tn * (i - 1);
                                        }
                                    }
                                }
                                else
                                {
                                    if (shape.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                                    {
                                        shape.TextFrame2.TextRange.Font.Fill.Transparency = tn0 + tn * (i - 1);
                                    }
                                }
                            }
                            label1.Text = "从" + tn0 * 100 + "%到" + tn1 * 100 + "%";
                        }
                    }
                }
            }

            if (comboBox1.SelectedItem as string == "线条宽度递进")
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    label1.Text = "起始值,终止值";
                }
                else
                {
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    if (sel.HasChildShapeRange)
                    {
                        range = sel.ChildShapeRange;
                    }
                    int count = range.Count;
                    string[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
                    if (arr.Count() == 1)
                    {
                        float tn0 = float.Parse(arr[0]);
                        for (int i = 1; i <= count; i++)
                        {
                            PowerPoint.Shape shape = range[i];
                            if (shape.Type == Office.MsoShapeType.msoGroup)
                            {
                                foreach (PowerPoint.Shape item in shape.GroupItems)
                                {
                                    if (item.Line.Visible == Office.MsoTriState.msoFalse)
                                    {
                                        item.Line.Visible = Office.MsoTriState.msoTrue;
                                    }
                                    item.Line.Weight = tn0;
                                }
                                label1.Text = "线宽是" + tn0 + "磅";
                            }
                            else
                            {
                                if (shape.Line.Visible == Office.MsoTriState.msoFalse)
                                {
                                    shape.Line.Visible = Office.MsoTriState.msoTrue;
                                }
                                shape.Line.Weight = tn0;
                                label1.Text = "线宽是" + tn0 + "磅";
                            }
                        }
                    }
                    else
                    {
                        float tn0 = float.Parse(arr[0]);
                        float tn1 = float.Parse(arr[1]);
                        if (count == 1 && range[1].Type == Office.MsoShapeType.msoGroup)
                        {
                            float tn = (tn1 - tn0) / (range[1].GroupItems.Count - 1);
                            for (int i = 1; i <= range[1].GroupItems.Count; i++)
                            {
                                PowerPoint.Shape shape = range[1].GroupItems[i];
                                if (shape.Line.Visible == Office.MsoTriState.msoFalse)
                                {
                                    shape.Line.Visible = Office.MsoTriState.msoTrue;
                                }
                                shape.Line.Weight = tn0 + tn * (i - 1);
                            }
                            label1.Text = "从" + tn0 + "到" + tn1 + "磅";
                        }
                        else
                        {
                            float tn = (tn1 - tn0) / (count - 1);
                            for (int i = 1; i <= count; i++)
                            {
                                PowerPoint.Shape shape = range[i];
                                if (shape.Type == Office.MsoShapeType.msoGroup)
                                {
                                    foreach (PowerPoint.Shape item in shape.GroupItems)
                                    {
                                        if (item.Line.Visible == Office.MsoTriState.msoFalse)
                                        {
                                            item.Line.Visible = Office.MsoTriState.msoTrue;
                                        }
                                        item.Line.Weight = tn0 + tn * (i - 1);
                                    }
                                    label1.Text = "从" + tn0+ "到" + tn1 + "磅";
                                }
                                else
                                {
                                    if (shape.Line.Visible == Office.MsoTriState.msoFalse)
                                    {
                                        shape.Line.Visible = Office.MsoTriState.msoTrue;
                                    }
                                    shape.Line.Weight = tn0 + tn * (i - 1);
                                    label1.Text = "从" + tn0 + "到" + tn1 + "磅";
                                }
                            }
                        }
                    }
                }
            }

            if (comboBox1.SelectedItem as string == "图片亮度递进")
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    label1.Text = "起始值,终止值";
                }
                else
                {
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    int count = range.Count;
                    string[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
                    if (arr.Count() == 1)
                    {
                        float bn0 = float.Parse(arr[0]) / 100;
                        for (int i = 1; i <= range.Count; i++)
                        {
                            PowerPoint.Shape shape = range[i];
                            if (shape.Type == Office.MsoShapeType.msoPicture)
                            {
                                if (shape.Fill.PictureEffects.Count != 0)
                                {
                                    int n = -1;
                                    for (int j = 1; j <= shape.Fill.PictureEffects.Count; j++)
                                    {
                                        if (shape.Fill.PictureEffects[j].Type == Office.MsoPictureEffectType.msoEffectBrightnessContrast)
                                        {
                                            Office.PictureEffect pics = shape.Fill.PictureEffects[j];
                                            pics.EffectParameters[1].Value = bn0;
                                            n = 1;
                                        }
                                    }
                                    if (n == -1)
                                    {
                                        Office.PictureEffect pics = shape.Fill.PictureEffects.Insert(Office.MsoPictureEffectType.msoEffectBrightnessContrast, 0);
                                        pics.EffectParameters[1].Value = bn0;
                                    }
                                }
                                else
                                {
                                    Office.PictureEffect pics = shape.Fill.PictureEffects.Insert(Office.MsoPictureEffectType.msoEffectBrightnessContrast, 0);
                                    pics.EffectParameters[1].Value = bn0;
                                    label1.Text = "图片亮度是" + bn0 * 100 + "%";
                                }
                            }
                            else
                            {
                                label1.Text = "需选中图片";
                            }
                            label1.Text = "图片亮度是" + bn0 * 100 + "%";
                        }
                    }
                    else
                    {
                        float bn0 = float.Parse(arr[0]) / 100;
                        float bn1 = float.Parse(arr[1]) / 100;
                        float bn = (bn1 - bn0) / (count - 1);
                        for (int i = 1; i <= count; i++)
                        {
                            PowerPoint.Shape shape = range[i];
                            if (shape.Type == Office.MsoShapeType.msoPicture)
                            {
                                if (shape.Fill.PictureEffects.Count != 0)
                                {
                                    int n = -1;
                                    for (int j = 1; j <= shape.Fill.PictureEffects.Count; j++)
			                        {
                                        if (shape.Fill.PictureEffects[j].Type == Office.MsoPictureEffectType.msoEffectBrightnessContrast)
	                                    {
                                            Office.PictureEffect pics = shape.Fill.PictureEffects[j];
                                            pics.EffectParameters[1].Value = bn0 + bn * (i - 1);
                                            n = 1;
	                                    }
			                        }
                                    if (n == -1)
                                    {
                                        Office.PictureEffect pics = shape.Fill.PictureEffects.Insert(Office.MsoPictureEffectType.msoEffectBrightnessContrast, 0);
                                        pics.EffectParameters[1].Value = bn0 + bn * (i - 1);
                                    }
                                }
                                else
                                {
                                    Office.PictureEffect pics = shape.Fill.PictureEffects.Insert(Office.MsoPictureEffectType.msoEffectBrightnessContrast, 0);
                                    pics.EffectParameters[1].Value = bn0 + bn * (i - 1);
                                    label1.Text = "从" + bn0 * 100 + "%到" + bn1 * 100 + "%";
                                }
                            }
                            else
                            {
                                label1.Text = "需选中图片";
                            }
                         }
                        label1.Text = "从" + bn0 * 100 + "%到" + bn1 * 100 + "%";
                    }
                }  
            }

            if (comboBox1.SelectedItem as string == "图片虚化递进")
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    label1.Text = "起始值,终止值";
                }
                else
                {
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    int count = range.Count;
                    string[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
                    if (arr.Count() == 1)
                    {
                        float bn0 = float.Parse(arr[0]);
                        for (int i = 1; i <= count; i++)
                        {
                            PowerPoint.Shape pic0 = range[i];
                            if (pic0.Type != Office.MsoShapeType.msoPicture)
                            {
                                label1.Text = "需选中图片";
                            }
                            else
                            {
                                if (pic0.Fill.PictureEffects.Count == 0)
                                {
                                    Office.PictureEffect piceff = pic0.Fill.PictureEffects.Insert(Office.MsoPictureEffectType.msoEffectBlur, 0);
                                    pic0.Fill.PictureEffects[1].EffectParameters[1].Value = bn0;
                                }
                                else
                                {
                                    int n = -1;
                                    for (int j = 1; j <= pic0.Fill.PictureEffects.Count; j++)
                                    {
                                        if (pic0.Fill.PictureEffects[j].Type == Office.MsoPictureEffectType.msoEffectBlur)
                                        {
                                            Office.PictureEffect piceff = pic0.Fill.PictureEffects[j];
                                            piceff.EffectParameters[1].Value = bn0;
                                            n = 1;
                                        }
                                    }
                                    if (n == -1)
                                    {
                                        Office.PictureEffect piceff = pic0.Fill.PictureEffects.Insert(Office.MsoPictureEffectType.msoEffectBlur, 0);
                                        pic0.Fill.PictureEffects[1].EffectParameters[1].Value = bn0;
                                    }
                                }
                            }
                        }
                        label1.Text = "图片虚化是" + bn0;
                    }
                    else
                    {
                        float bn0 = float.Parse(arr[0]);
                        float bn1 = float.Parse(arr[1]);
                        float bn = (bn1 - bn0) / (count - 1);
                        for (int i = 1; i <= count; i++)
                        {
                            PowerPoint.Shape pic0 = range[i];
                            if (pic0.Type != Office.MsoShapeType.msoPicture)
                            {
                                label1.Text = "需选中图片";
                            }
                            else
                            {
                                if (pic0.Fill.PictureEffects.Count == 0)
                                {
                                    Office.PictureEffect piceff = pic0.Fill.PictureEffects.Insert(Office.MsoPictureEffectType.msoEffectBlur, 0);
                                    pic0.Fill.PictureEffects[1].EffectParameters[1].Value = bn0 + bn * (i - 1);
                                }
                                else
                                {
                                    int n = -1;
                                    for (int j = 1; j <= pic0.Fill.PictureEffects.Count; j++)
                                    {
                                        if (pic0.Fill.PictureEffects[j].Type == Office.MsoPictureEffectType.msoEffectBlur)
                                        {
                                            Office.PictureEffect piceff = pic0.Fill.PictureEffects[j];
                                            pic0.Fill.PictureEffects[1].EffectParameters[1].Value = bn0 + bn * (i - 1);
                                            n = 1;
                                        }
                                    }
                                    if (n == -1)
                                    {
                                        Office.PictureEffect piceff = pic0.Fill.PictureEffects.Insert(Office.MsoPictureEffectType.msoEffectBlur, 0);
                                        pic0.Fill.PictureEffects[1].EffectParameters[1].Value = bn0 + bn * (i - 1);
                                    }
                                }
                            }    
                        }
                        label1.Text = "虚化从" + bn0 + "到" + bn1;
                    }
                }
            }

            if (comboBox1.SelectedItem as string == "距底边高度递进")
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    label1.Text = "起始值,终止值";
                }
                else
                {
                    PowerPoint.ShapeRange range = sel.ShapeRange;
                    if (sel.HasChildShapeRange)
                    {
                        range = sel.ChildShapeRange;
                    }
                    int count = range.Count;
                    String[] arr = textBox1.Text.Trim().Split(char.Parse(","), char.Parse(" "), char.Parse("，")).ToArray();
                    if (arr.Count() == 1)
                    {
                        float bn0 = float.Parse(arr[0]);
                        if (count == 1 && range[1].Type == Office.MsoShapeType.msoGroup)
                        {
                            foreach (PowerPoint.Shape item in range[1].GroupItems)
                            {
                                item.ThreeD.Z = bn0;
                            }
                        }
                        else
                        {
                            for (int i = 1; i <= count; i++)
                            {
                                range[i].ThreeD.Z = bn0;
                            }
                        }
                        label1.Text = "高度值:" + bn0;
                    }
                    else
                    {
                        float bn0 = float.Parse(arr[0]);
                        float bn1 = float.Parse(arr[1]);
                        if (count == 1 && range[1].Type == Office.MsoShapeType.msoGroup)
                        {
                            for (int i = 1; i <= range[1].GroupItems.Count; i++)
                            {
                                float bn = (bn1 - bn0) / (range[1].GroupItems.Count - 1);
                                range[1].GroupItems[i].ThreeD.Z = bn0 + bn * (i - 1);
                            }
                        }
                        else
                        {
                            for (int j = 1; j <= count; j++)
                            {
                                float bn = (bn1 - bn0) / (count - 1);
                                range[j].ThreeD.Z = bn0 + bn * (j - 1);
                            }
                        }
                        label1.Text = "从" + bn0 + "到" + bn1;
                    }
                } 
            }
        }

        private void fast1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "jpg(*.jpg)|*.jpg|All files (*.*)|*.*";
                openFileDialog.RestoreDirectory = true;
                openFileDialog.FilterIndex = 1;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string fName = openFileDialog.FileName;
                    this.BackgroundImage = Image.FromFile(fName);
                }
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == System.Convert.ToChar(13))
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                PowerPoint.Selection sel = app.ActiveWindow.Selection;
                if (comboBox1.SelectedItem as string == "填充色等值递进(回车)")
                {
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                    {
                        label1.Text = "色型,递进值";
                    }
                    else
                    {
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        if (sel.HasChildShapeRange)
                        {
                            range = sel.ChildShapeRange;
                        }
                        int count = range.Count;

                        for (int i = 1; i <= count; i++)
                        {
                            PowerPoint.Shape shape = range[i];
                            if (shape.Type == Office.MsoShapeType.msoGroup)
                            {
                                int count2 = shape.GroupItems.Count;
                                for (int j = 1; j <= count2; j++)
                                {
                                    PowerPoint.Shape gshape = shape.GroupItems[j];
                                    if (gshape.Shadow.Visible == Office.MsoTriState.msoTrue)
                                    {
                                        float tr = gshape.Shadow.Transparency;
                                        int rgb0 = gshape.Shadow.ForeColor.RGB;
                                        int nrgb = Nrgb(rgb0);
                                        gshape.Shadow.ForeColor.RGB = nrgb;
                                        gshape.Shadow.Transparency = tr;
                                    }
                                    if (gshape.Fill.Type == Office.MsoFillType.msoFillGradient)
                                    {
                                        int count3 = gshape.Fill.GradientStops.Count;
                                        for (int k = 1; k <= count3; k++)
                                        {
                                            int rgb0 = gshape.Fill.GradientStops[k].Color.RGB;
                                            float tr = gshape.Fill.GradientStops[k].Transparency;
                                            int nrgb = Nrgb(rgb0);
                                            gshape.Fill.GradientStops[k].Color.RGB = nrgb;
                                            gshape.Fill.GradientStops[k].Transparency = tr;
                                        }
                                    }
                                    else if (gshape.Fill.Type == Office.MsoFillType.msoFillSolid)
                                    {
                                        float tr = gshape.Fill.Transparency;
                                        int rgb0 = gshape.Fill.ForeColor.RGB;
                                        int nrgb = Nrgb(rgb0);
                                        gshape.Fill.ForeColor.RGB = nrgb;
                                        gshape.Fill.Transparency = tr;
                                    }
                                    else
                                    {
                                        label1.Text = "仅支持纯色和渐变";
                                    }
                                }
                            }
                            else
                            {
                                if (shape.Shadow.Visible == Office.MsoTriState.msoTrue)
                                {
                                    float tr = shape.Shadow.Transparency;
                                    int rgb0 = shape.Shadow.ForeColor.RGB;
                                    int nrgb = Nrgb(rgb0);
                                    shape.Shadow.ForeColor.RGB = nrgb;
                                    shape.Shadow.Transparency = tr;
                                }
                                if (shape.Fill.Type == Office.MsoFillType.msoFillGradient)
                                {
                                    int count4 = shape.Fill.GradientStops.Count;
                                    for (int m = 1; m <= count4; m++)
                                    {
                                        int rgb0 = shape.Fill.GradientStops[m].Color.RGB;
                                        float tr = shape.Fill.GradientStops[m].Transparency;
                                        int nrgb = Nrgb(rgb0);
                                        shape.Fill.GradientStops[m].Color.RGB = nrgb;
                                        shape.Fill.GradientStops[m].Transparency = tr;
                                    }
                                }
                                else if (shape.Fill.Type == Office.MsoFillType.msoFillSolid)
                                {
                                    float tr = shape.Fill.Transparency;
                                    int rgb0 = shape.Fill.ForeColor.RGB;
                                    int nrgb = Nrgb(rgb0);
                                    shape.Fill.ForeColor.RGB = nrgb;
                                    shape.Fill.Transparency = tr;
                                }
                                else
                                {
                                    label1.Text = "仅支持纯色和渐变";
                                }
                            }
                        }
                    }
                }

                if (comboBox1.SelectedItem as string == "线条色等值递进(回车)")
                {
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                    {
                        label1.Text = "色型,递进值";
                    }
                    else
                    {
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        if (sel.HasChildShapeRange)
                        {
                            range = sel.ChildShapeRange;
                        }
                        int count = range.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            PowerPoint.Shape shape = range[i];
                            if (shape.Type == Office.MsoShapeType.msoGroup)
                            {
                                int count2 = shape.GroupItems.Count;
                                for (int j = 1; j <= count2; j++)
                                {
                                    PowerPoint.Shape gshape = shape.GroupItems[j];
                                    if (gshape.Line.Visible == Office.MsoTriState.msoTrue)
                                    {
                                        float tr = gshape.Line.Transparency;
                                        int rgb0 = gshape.Line.ForeColor.RGB;
                                        int nrgb = Nrgb(rgb0);
                                        gshape.Line.ForeColor.RGB = nrgb;
                                        gshape.Line.Transparency = tr;
                                    }
                                    else
                                    {
                                        label1.Text = "先添加线条色";
                                    }
                                }
                            }
                            else
                            {
                                if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                                {
                                    float tr = shape.Line.Transparency;
                                    int rgb0 = shape.Line.ForeColor.RGB;
                                    int nrgb = Nrgb(rgb0);
                                    shape.Line.ForeColor.RGB = nrgb;
                                    shape.Line.Transparency = tr;
                                }
                                else
                                {
                                    label1.Text = "先添加线条色";
                                }
                            }
                        }
                    }
                }

                if (comboBox1.SelectedItem as string == "阴影色等值递进(回车)")
                {
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                    {
                        label1.Text = "色型,递进值";
                    }
                    else
                    {
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        if (sel.HasChildShapeRange)
                        {
                            range = sel.ChildShapeRange;
                        }
                        int count = range.Count;

                        for (int i = 1; i <= count; i++)
                        {
                            PowerPoint.Shape shape = range[i];
                            if (shape.Type == Office.MsoShapeType.msoGroup)
                            {
                                int count2 = shape.GroupItems.Count;
                                for (int j = 1; j <= count2; j++)
                                {
                                    PowerPoint.Shape gshape = shape.GroupItems[j];
                                    if (gshape.Shadow.Visible == Office.MsoTriState.msoTrue)
                                    {
                                        float tr = gshape.Shadow.Transparency;
                                        int rgb0 = gshape.Shadow.ForeColor.RGB;
                                        int nrgb = Nrgb(rgb0);
                                        gshape.Shadow.ForeColor.RGB = nrgb;
                                        gshape.Shadow.Transparency = tr;
                                    }
                                    else
                                    {
                                        label1.Text = "先设置阴影";
                                    }
                                }
                            }
                            else
                            {
                                if (shape.Shadow.Visible == Office.MsoTriState.msoTrue)
                                {
                                    float tr = shape.Shadow.Transparency;
                                    int rgb0 = shape.Shadow.ForeColor.RGB;
                                    int nrgb = Nrgb(rgb0);
                                    shape.Shadow.ForeColor.RGB = nrgb;
                                    shape.Shadow.Transparency = tr;
                                }
                                else
                                {
                                    label1.Text = "先设置阴影";
                                }
                            }
                        }
                    }
                }

                if (comboBox1.SelectedItem as string == "填充色差值递进(回车)")
                {
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                    {
                        label1.Text = "色型,递进值";
                    }
                    else
                    {
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        if (sel.HasChildShapeRange)
                        {
                            range = sel.ChildShapeRange;
                        }
                        int count = range.Count;

                        for (int i = 1; i <= count; i++)
                        {
                            PowerPoint.Shape shape = range[i];
                            if (shape.Type == Office.MsoShapeType.msoGroup)
                            {
                                int count2 = shape.GroupItems.Count;
                                for (int j = 1; j <= count2; j++)
                                {
                                    PowerPoint.Shape gshape = shape.GroupItems[j];
                                    if (gshape.Fill.Type == Office.MsoFillType.msoFillGradient)
                                    {
                                        int count3 = gshape.Fill.GradientStops.Count;
                                        for (int k = 1; k <= count3; k++)
                                        {
                                            float tr = gshape.Fill.GradientStops[k].Transparency;
                                            int rgb0 = gshape.Fill.GradientStops[k].Color.RGB;
                                            int nrgb = -1;
                                            if (count == 1)
                                            {
                                                nrgb = Nrgb2(rgb0, j);
                                            }
                                            else
                                            {
                                                nrgb = Nrgb2(rgb0, i);
                                            }
                                            gshape.Fill.GradientStops[k].Color.RGB = nrgb;
                                            gshape.Fill.GradientStops[k].Transparency = tr;
                                        }
                                    }
                                    else if (gshape.Fill.Type == Office.MsoFillType.msoFillSolid)
                                    {
                                        int rgb0 = gshape.Fill.ForeColor.RGB;
                                        int nrgb = -1;
                                        if (count == 1)
                                        {
                                            nrgb = Nrgb2(rgb0, j);
                                        }
                                        else
                                        {
                                            nrgb = Nrgb2(rgb0, i);
                                        }
                                        gshape.Fill.ForeColor.RGB = nrgb;
                                    }
                                    else
                                    {
                                        label1.Text = "仅支持纯色和渐变";
                                    }
                                }
                            }
                            else
                            {
                                if (shape.Fill.Type == Office.MsoFillType.msoFillGradient)
                                {
                                    int count4 = shape.Fill.GradientStops.Count;
                                    for (int m = 1; m <= count4; m++)
                                    {
                                        float tr = shape.Fill.GradientStops[m].Transparency;
                                        int rgb0 = shape.Fill.GradientStops[m].Color.RGB;
                                        int nrgb = Nrgb2(rgb0, i);
                                        shape.Fill.GradientStops[m].Color.RGB = nrgb;
                                        shape.Fill.GradientStops[m].Transparency = tr;
                                    }
                                }
                                else if (shape.Fill.Type == Office.MsoFillType.msoFillSolid)
                                {
                                    int rgb0 = shape.Fill.ForeColor.RGB;
                                    int nrgb = Nrgb2(rgb0, i);
                                    shape.Fill.ForeColor.RGB = nrgb;
                                }
                                else
                                {
                                    label1.Text = "仅支持纯色和渐变";
                                }
                            }
                        }
                    }
                }

                if (comboBox1.SelectedItem as string == "线条色差值递进(回车)")
                {
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                    {
                        label1.Text = "色型,递进值";
                    }
                    else
                    {
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        if (sel.HasChildShapeRange)
                        {
                            range = sel.ChildShapeRange;
                        }
                        int count = range.Count;

                        for (int i = 1; i <= count; i++)
                        {
                            PowerPoint.Shape shape = range[i];
                            if (shape.Type == Office.MsoShapeType.msoGroup)
                            {
                                int count2 = shape.GroupItems.Count;
                                for (int j = 1; j <= count2; j++)
                                {
                                    PowerPoint.Shape gshape = shape.GroupItems[j];
                                    if (gshape.Line.Visible == Office.MsoTriState.msoTrue)
                                    {
                                        float tr = gshape.Line.Transparency;
                                        int rgb0 = gshape.Line.ForeColor.RGB;
                                        int nrgb = -1;
                                        if (count == 1)
                                        {
                                            nrgb = Nrgb2(rgb0, j);
                                        }
                                        else
                                        {
                                            nrgb = Nrgb2(rgb0, i);
                                        }
                                        gshape.Line.ForeColor.RGB = nrgb;
                                        gshape.Line.Transparency = tr;
                                    }
                                    else
                                    {
                                        label1.Text = "先添加线条色";
                                    }
                                }
                            }
                            else
                            {
                                if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                                {
                                    float tr = shape.Line.Transparency;
                                    int rgb0 = shape.Line.ForeColor.RGB;
                                    int nrgb = Nrgb2(rgb0, i);
                                    shape.Line.ForeColor.RGB = nrgb;
                                    shape.Line.Transparency = tr;
                                }
                                else
                                {
                                    label1.Text = "先添加线条色";
                                }
                            }
                        }
                    }
                }

                if (comboBox1.SelectedItem as string == "阴影色差值递进(回车)")
                {
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                    {
                        label1.Text = "色型,递进值";
                    }
                    else
                    {
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        if (sel.HasChildShapeRange)
                        {
                            range = sel.ChildShapeRange;
                        }
                        int count = range.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            PowerPoint.Shape shape = range[i];
                            if (shape.Type == Office.MsoShapeType.msoGroup)
                            {
                                int count2 = shape.GroupItems.Count;
                                for (int j = 1; j <= count2; j++)
                                {
                                    PowerPoint.Shape gshape = shape.GroupItems[j];
                                    if (gshape.Shadow.Visible == Office.MsoTriState.msoTrue)
                                    {
                                        float tr = gshape.Shadow.Transparency;
                                        int rgb0 = gshape.Shadow.ForeColor.RGB;
                                        int nrgb = -1;
                                        if (count == 1)
                                        {
                                            nrgb = Nrgb2(rgb0, j);
                                        }
                                        else
                                        {
                                            nrgb = Nrgb2(rgb0, i);
                                        }
                                        gshape.Shadow.ForeColor.RGB = nrgb;
                                        gshape.Shadow.Transparency = tr;
                                    }
                                    else
                                    {
                                        label1.Text = "先设置阴影";
                                    }
                                }
                            }
                            else
                            {
                                if (shape.Shadow.Visible == Office.MsoTriState.msoTrue)
                                {
                                    float tr = shape.Shadow.Transparency;
                                    int rgb0 = shape.Shadow.ForeColor.RGB;
                                    int nrgb = Nrgb2(rgb0, i);
                                    shape.Shadow.ForeColor.RGB = nrgb;
                                    shape.Shadow.Transparency = tr;
                                }
                                else
                                {
                                    label1.Text = "先设置阴影";
                                }
                            }
                        }
                    }
                }

                if (comboBox1.SelectedItem as string == "三维旋转递进(回车)")
                {
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                    {
                        label1.Text = "类型,初始值,终止值";
                    }
                    else
                    {
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        int count = range.Count;
                        string[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
                        if (arr.Count() == 2)
                        {
                            string type= arr[0].ToString();
                            float dj = float.Parse(arr[1]);
                            for (int i = 1; i <= count; i++)
                            {
                                PowerPoint.Shape shape = range[i];
                                shape.ThreeD.Visible = Office.MsoTriState.msoTrue;
                                if (type == "x" || type == "X")
                                {
                                    shape.ThreeD.RotationX += 360 - dj;
                                }
                                if (type == "y" || type == "Y")
                                {
                                    shape.ThreeD.RotationY += dj;
                                }
                                if (type == "z" || type == "Z")
                                {
                                    shape.ThreeD.RotationZ += 360 - dj;
                                }
                                if (type == "xy" || type == "XY")
                                {
                                    shape.ThreeD.RotationX += 360 - dj;
                                    shape.ThreeD.RotationY += dj;
                                }
                                if (type == "xz" || type == "XZ")
                                {
                                    shape.ThreeD.RotationX += 360 - dj;
                                    shape.ThreeD.RotationZ += 360 - dj;
                                }
                                if (type == "yz" || type == "YZ")
                                {
                                    shape.ThreeD.RotationY += dj;
                                    shape.ThreeD.RotationZ += 360 - dj;
                                }
                                if (type == "xyz" || type == "XYZ")
                                {
                                    shape.ThreeD.RotationX += 360 - dj;
                                    shape.ThreeD.RotationY += dj;
                                    shape.ThreeD.RotationZ += 360 - dj;
                                }
                                label1.Text = type + "递进" + dj;
                            }
                        }
                        else if (arr.Count() == 3)
                        {
                            string type= arr[0].ToString();
                            float dj0 = float.Parse(arr[1]);
                            float dj1 = float.Parse(arr[2]);
                            float n = 0;
                            if (count == 1)
                            {
                                n = dj1;
                            }
                            else if (count > 1)
                            {
                                n = (dj1 - dj0) / (count - 1);
                            }
                            else
                            {
                                label1.Text = "先选中形状";
                            }
                            
                            for (int i = 1; i <= count; i++)
                            {
                                PowerPoint.Shape shape = range[i];
                                shape.ThreeD.Visible = Office.MsoTriState.msoTrue;
                                float dj = dj0 + n * (i - 1);
                                if (type == "x" || type == "X")
                                {
                                    shape.ThreeD.RotationX = 360 - dj;
                                }
                                if (type == "y" || type == "Y")
                                {
                                    shape.ThreeD.RotationY = dj;
                                }
                                if (type == "z" || type == "Z")
                                {
                                    shape.ThreeD.RotationZ = 360 - dj;
                                }
                                if (type == "xy" || type == "XY")
                                {
                                    shape.ThreeD.RotationX = 360 - dj;
                                    shape.ThreeD.RotationY = dj;
                                }
                                if (type == "xz" || type == "XZ")
                                {
                                    shape.ThreeD.RotationX = 360 - dj;
                                    shape.ThreeD.RotationZ = 360 - dj;
                                }
                                if (type == "yz" || type == "YZ")
                                {
                                    shape.ThreeD.RotationY = dj;
                                    shape.ThreeD.RotationZ = 360 - dj;
                                }
                                if (type == "xyz" || type == "XYZ")
                                {
                                    shape.ThreeD.RotationX = 360 - dj;
                                    shape.ThreeD.RotationY = dj;
                                    shape.ThreeD.RotationZ = 360 - dj;
                                }
                                label1.Text = type + "从" + dj0 + "到" + dj1;
                            }
                        }
                    }
                }

                if (comboBox1.SelectedItem as string == "图片分割(回车)")
                {
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                    {
                        label1.Text = "行数,列数";
                    }
                    else
                    {
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        int count = range.Count;
                        string[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
                        int n1 = int.Parse(arr[0]);
                        int n2 = int.Parse(arr[1]);
                        string apath = app.ActivePresentation.Path;
                        for (int i = 1; i <= count; i++)
                        {
                            PowerPoint.Shape pic = range[i];
                            pic.Copy();
                            PowerPoint.Shape npic = slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
                            npic.Left = pic.Left + pic.Width / 2 - npic.Width / 2;
                            npic.Top = pic.Top + pic.Height / 2 - npic.Height / 2;
                            for (int j = 1; j <= n1; j++)
                            {
                                for (int k = 1; k <= n2; k++)
                                {
                                    PowerPoint.Shape nnpic = npic.Duplicate()[1];
                                    nnpic.PictureFormat.CropTop = npic.Height / n1 * (j - 1);
                                    nnpic.PictureFormat.CropBottom = npic.Height / n1 * (n1 - j);
                                    nnpic.PictureFormat.CropLeft = npic.Width / n2 * (k - 1);
                                    nnpic.PictureFormat.CropRight = npic.Width / n2 * (n2 - k);
                                    nnpic.Top = npic.Top + npic.Height / n1 * (j - 1);
                                    nnpic.Left = npic.Left + npic.Width / n2 * (k - 1);
                                }
                            }
                            npic.Delete();
                            pic.Delete();
                            label1.Text = n1 + "行," + n2 + "列";
                        }
                    }
                }

                if (comboBox1.SelectedItem as string == "图片马赛克(回车)")
                {
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                    {
                        label1.Text = "马赛克大小";
                    }
                    else
                    {
                        string[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
                        int nc = int.Parse(arr[0]);
                        string apath = app.ActivePresentation.Path;
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        int count = range.Count;
                        for (int p = 1; p <= count; p++)
                        {
                            PowerPoint.Shape npic = range[p];
                            npic.Copy();
                            PowerPoint.Shape pic = slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
                            pic.Export(apath + @"xshape.png", PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                            Bitmap bmp = new Bitmap(apath + @"xshape.png");
                            Graphics g = Graphics.FromImage(bmp);
                            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                            g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                            g.DrawImage(bmp, 0, 0);
                            g.Dispose();

                            for (int i = 0; i < bmp.Width / nc; i++)
                            {
                                for (int j = 0; j < bmp.Height / nc; j++)
                                {
                                    Color color = bmp.GetPixel(i * nc, j * nc);
                                    for (int m = 0; m < nc; m++)
                                    {
                                        for (int n = 0; n < nc; n++)
                                        {
                                            bmp.SetPixel(i * nc + m, j * nc + n, color);
                                        }
                                    }

                                }
                            }
                            bmp.Save(apath + @"xshape2.png");
                            PowerPoint.Shape nshape = slide.Shapes.AddPicture(apath + @"xshape2.png", Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, npic.Left, npic.Top, pic.Width, pic.Height);
                            pic.Delete();
                            npic.Delete();
                            bmp.Dispose();
                            File.Delete(apath + @"xshape.png");
                            File.Delete(apath + @"xshape2.png");
                            label1.Text = "马赛克:" + nc;
                        } 
                    }
                }

                if (comboBox1.SelectedItem as string == "图片色相(回车)")
                {
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                    {
                        label1.Text = "色调值";
                    }
                    else
                    {
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        string apath = app.ActivePresentation.Path;
                        int count = range.Count;
                        string[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
                        if (arr.Count() == 1)
                        {
                            int h = int.Parse(arr[0]);
                            for (int p = 1; p <= count; p++)
                            {
                                PowerPoint.Shape npic = range[p];
                                npic.Copy();
                                PowerPoint.Shape pic = slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
                                pic.Export(apath + @"xshape.png", PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                                Bitmap bmp = new Bitmap(apath + @"xshape.png");
                                Graphics g = Graphics.FromImage(bmp);
                                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                                g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                                g.DrawImage(bmp, 0, 0);
                                g.Dispose();

                                for (int i = 0; i < bmp.Width; i++)
                                {
                                    for (int j = 0; j < bmp.Height; j++)
                                    {
                                        int na, nr, ng, nb = 0;
                                        Color color = bmp.GetPixel(i, j);
                                        na = color.A;
                                        if (na != 0)
                                        {
                                            nr = color.R;
                                            ng = color.G;
                                            nb = color.B;
                                            int hsl = Rgb2Hsl(nr, ng, nb);
                                            int s = (hsl / 256) % 256;
                                            int l = (hsl / 256 / 256) % 256;
                                            int nrgb = Hsl2Rgb(h, s, l);
                                            nr = nrgb % 256;
                                            ng = (nrgb / 256) % 256;
                                            nb = (nrgb / 256 / 256) % 256;
                                            bmp.SetPixel(i, j, Color.FromArgb(na, nr, ng, nb));
                                        }
                                    }
                                }
                                bmp.Save(apath + @"xshape2.png");
                                PowerPoint.Shape nshape = slide.Shapes.AddPicture(apath + @"xshape2.png", Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, npic.Left + npic.Width, npic.Top + npic.Height / 2 - pic.Height / 2, pic.Width, pic.Height);
                                pic.Delete();
                                bmp.Dispose();
                                File.Delete(apath + @"xshape.png");
                                File.Delete(apath + @"xshape2.png");
                                label1.Text = "色相为" + h;
                            }
                        }
                    }
                }

                if (comboBox1.SelectedItem as string == "数字递进(回车)")
                {
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                    {
                        label1.Text = "起始值,递进值";
                    }
                    else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        if (sel.HasChildShapeRange)
                        {
                            range = sel.ChildShapeRange;
                        }
                        string[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
                        if (arr.Count() == 2)
                        {
                            float n1 = float.Parse(arr[0]);
                            float n2 = float.Parse(arr[1]);
                            int count = range.Count;
                            for (int i = 1; i <= count; i++)
                            {
                                PowerPoint.Shape shape = range[i];
                                shape.TextFrame2.TextRange.Characters.Text = Convert.ToString(n1 + n2 * (i - 1));
                            }
                            label1.Text = "已递进";
                        }
                        else if (arr.Count() == 3)
                        {
                            float n1 = float.Parse(arr[1]);
                            float n2 = float.Parse(arr[2]);
                            int count = range.Count;
                            for (int i = 1; i <= count; i++)
                            {
                                PowerPoint.Shape shape = range[i];
                                shape.TextFrame2.TextRange.Characters.Text = arr[0] + Convert.ToString(n1 + n2 * (i - 1));
                            }
                            label1.Text = "已递进";
                        }
                        else if (arr.Count() == 4)
                        {
                            float n1 = float.Parse(arr[1]);
                            float n2 = float.Parse(arr[2]);
                            int count = range.Count;
                            for (int i = 1; i <= count; i++)
                            {
                                PowerPoint.Shape shape = range[i];
                                shape.TextFrame2.TextRange.Characters.Text = arr[0] + Convert.ToString(n1 + n2 * (i - 1)) + arr[3];
                            }
                            label1.Text = "已递进";
                        }
                    }
                    else
                    {
                        label1.Text = "请选中形状";
                    }
                }

                if (comboBox1.SelectedItem as string == "跨页复制(回车)")
                {
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                    {
                        label1.Text = "起始页,终止页";
                    }
                    else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        PowerPoint.Slides slides = app.ActivePresentation.Slides;
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        if (sel.HasChildShapeRange)
                        {
                            range = sel.ChildShapeRange;
                        }
                        string[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
                        if (arr.Count() != 2)
                        {
                            label1.Text = "例如 1,3";
                        }
                        else
                        {
                            int n1 = int.Parse(arr[0]);
                            int n2 = int.Parse(arr[1]);
                            if (n2 <= slides.Count && n1 > 0)
                            {
                                int count = range.Count;
                                for (int i = 1; i <= count; i++)
                                {
                                    PowerPoint.Shape shape = range[i];
                                    shape.Copy();
                                    for (int j = n1; j <= n2; j++)
                                    {
                                        if (j != slide.SlideNumber)
                                        {
                                            PowerPoint.Shape nshape = slides[j].Shapes.Paste()[1];
                                            nshape.Left = shape.Left;
                                            nshape.Top = shape.Top;
                                        }
                                    }
                                }
                                label1.Text = "已复制到" + n1 + "-" + n2 + "页";
                            }
                            else
                            {
                                label1.Text = "超出页数";
                            }   
                        }
                    }
                    else
                    {
                        label1.Text = "请选中形状";
                    }
                }

                if (comboBox1.SelectedItem as string == "尺寸比例递进(回车)")
                {
                    if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        MessageBox.Show("请先选中至少一个图形");
                    }
                    else
                    {
                        string[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        if (sel.HasChildShapeRange)
                        {
                            range = sel.ChildShapeRange;
                        }
                        int count = range.Count;
                        float w, h, l, t;
                        if (arr.Count() <= 1)
                        {
                            float n = float.Parse(arr[0]) * 0.01f;
                            if (count == 1 && range.Type == Office.MsoShapeType.msoGroup)
                            {
                                for (int i = 1; i <= range.GroupItems.Count; i++)
                                {
                                    w = range[1].GroupItems[i].Width;
                                    h = range[1].GroupItems[i].Height;
                                    l = range[1].GroupItems[i].Left+w/2;
                                    t = range[1].GroupItems[i].Top+h/2;
                                    w = w * n;
                                    h = h * n;
                                    range[1].GroupItems[i].Width = w;
                                    range[1].GroupItems[i].Height = h;
                                    range[1].GroupItems[i].Left = l - w / 2;
                                    range[1].GroupItems[i].Top = t - h / 2;
                                }
                            }
                            else
                            {
                                for (int i = 1; i <= count; i++)
                                {
                                    w = range[i].Width;
                                    h = range[i].Height;
                                    l = range[i].Left + w / 2;
                                    t = range[i].Top + h / 2;
                                    w = w * n;
                                    h = h * n;
                                    range[i].Width = w;
                                    range[i].Height = h;
                                    range[i].Left = l - w / 2;
                                    range[i].Top = t - h / 2;
                                }
                            }
                            label1.Text = "尺寸缩放：" + n * 100 + "%";
                        }
                        else
                        {
                            float n1 = float.Parse(arr[0]) * 0.01f;
                            float n2 = float.Parse(arr[1]) * 0.01f;
                            float nn;
                            if (count == 1 && range.Type == Office.MsoShapeType.msoGroup)
                            {
                                nn = (n2 - n1) / (range.GroupItems.Count - 1);
                                for (int i = 1; i <= range.GroupItems.Count; i++)
                                {
                                    w = range[1].GroupItems[i].Width;
                                    h = range[1].GroupItems[i].Height;
                                    l = range[1].GroupItems[i].Left + w / 2;
                                    t = range[1].GroupItems[i].Top + h / 2;
                                    w = w * (n1 + nn * (i - 1));
                                    h = h * (n1 + nn * (i - 1));
                                    range[1].GroupItems[i].Width = w;
                                    range[1].GroupItems[i].Height = h;
                                    range[1].GroupItems[i].Left = l - w / 2;
                                    range[1].GroupItems[i].Top = t - h / 2;
                                }
                            }
                            else
                            {
                                if (count == 1)
                                {
                                    nn = 0;
                                }
                                else
                                {
                                    nn = (n2 - n1) / (range.Count - 1);
                                }         
                                for (int i = 1; i <= count; i++)
                                {
                                    w = range[i].Width;
                                    h = range[i].Height;
                                    l = range[i].Left + w / 2;
                                    t = range[i].Top + h / 2;
                                    w = w * (n1 + nn * (i - 1));
                                    h = h * (n1 + nn * (i - 1));
                                    range[i].Width = w;
                                    range[i].Height = h;
                                    range[i].Left = l - w / 2;
                                    range[i].Top = t - h / 2;
                                }
                            }
                            label1.Text = "尺寸：" + n1 * 100 + "%到" + n2 * 100 + "%";
                        }                      
                    }  
                }

                if (comboBox1.SelectedItem as string == "高宽比例缩放(回车)")
                {
                    if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        MessageBox.Show("请先选中至少一个图形");
                    }
                    else
                    {
                        string[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        if (sel.HasChildShapeRange)
                        {
                            range = sel.ChildShapeRange;
                        }
                        int count = range.Count;
                        float w, h, l, t;
                        if (arr.Count() <= 1)
                        {
                            float n = float.Parse(arr[0]) * 0.01f;
                            if (count == 1 && range.Type == Office.MsoShapeType.msoGroup)
                            {
                                for (int i = 1; i <= range.GroupItems.Count; i++)
                                {
                                    w = range[1].GroupItems[i].Width;
                                    h = range[1].GroupItems[i].Height;
                                    l = range[1].GroupItems[i].Left+w/2;
                                    t = range[1].GroupItems[i].Top+h/2;
                                    w = w * n;
                                    h = h * n;
                                    range[1].GroupItems[i].Width = w;
                                    range[1].GroupItems[i].Height = h;
                                    range[1].GroupItems[i].Left = l - w / 2;
                                    range[1].GroupItems[i].Top = t - h / 2;
                                }
                            }
                            else
                            {
                                for (int i = 1; i <= count; i++)
                                {
                                    w = range[i].Width;
                                    h = range[i].Height;
                                    l = range[i].Left + w / 2;
                                    t = range[i].Top + h / 2;
                                    w = w * n;
                                    h = h * n;
                                    range[i].Width = w;
                                    range[i].Height = h;
                                    range[i].Left = l - w / 2;
                                    range[i].Top = t - h / 2;
                                }
                            }
                            label1.Text = "高宽缩放：" + n * 100 + "%";
                        }
                        else
                        {
                            float n1 = float.Parse(arr[0]) * 0.01f;
                            float n2 = float.Parse(arr[1]) * 0.01f;
                            if (count == 1 && range.Type == Office.MsoShapeType.msoGroup)
                            {
                                for (int i = 1; i <= range.GroupItems.Count; i++)
                                {
                                    if (range[1].GroupItems[i].LockAspectRatio == Office.MsoTriState.msoTrue)
                                    {
                                        range[1].GroupItems[i].LockAspectRatio = Office.MsoTriState.msoFalse;
                                    }
                                    w = range[1].GroupItems[i].Width;
                                    h = range[1].GroupItems[i].Height;
                                    l = range[1].GroupItems[i].Left + w / 2;
                                    t = range[1].GroupItems[i].Top + h / 2;
                                    w = w * n2;
                                    h = h * n1;
                                    range[1].GroupItems[i].Width = w;
                                    range[1].GroupItems[i].Height = h;
                                    range[1].GroupItems[i].Left = l - w / 2;
                                    range[1].GroupItems[i].Top = t - h / 2;
                                }
                            }
                            else
                            {
                                for (int i = 1; i <= count; i++)
                                {
                                    if (range[i].LockAspectRatio == Office.MsoTriState.msoTrue)
                                    {
                                        range[i].LockAspectRatio = Office.MsoTriState.msoFalse;
                                    }
                                    w = range[i].Width;
                                    h = range[i].Height;
                                    l = range[i].Left + w / 2;
                                    t = range[i].Top + h / 2;
                                    w = w * n2;
                                    h = h * n1;
                                    range[i].Width = w;
                                    range[i].Height = h;
                                    range[i].Left = l - w / 2;
                                    range[i].Top = t - h / 2;
                                }
                            }
                            label1.Text = "高：" + n1 * 100 + "%  宽：" + n2 * 100 + "%";
                        }
                    }  
                }

                if (comboBox1.SelectedItem as string == "图片比例裁剪(回车)")
                {
                    if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        MessageBox.Show("请先选中至少一个图片");
                    }
                    else
                    {
                        string[] arr = textBox1.Text.Trim().Split(char.Parse(":"), char.Parse("："), char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        if (sel.HasChildShapeRange)
                        {
                            range = sel.ChildShapeRange;
                        }
                        int count = range.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            PowerPoint.Shape pic = range[i];
                            if (pic.Type == Office.MsoShapeType.msoPicture)
                            {
                                float lm = pic.Left + pic.Width / 2;
                                float tm = pic.Top + pic.Height / 2;
                                if (arr.Count() <= 1 || arr[0] == "0")
                                {
                                    pic.PictureFormat.Crop.ShapeWidth = pic.PictureFormat.Crop.PictureWidth;
                                    pic.PictureFormat.Crop.ShapeHeight = pic.PictureFormat.Crop.PictureHeight;
                                    pic.PictureFormat.Crop.PictureOffsetX = (pic.PictureFormat.Crop.ShapeWidth - pic.PictureFormat.Crop.PictureWidth) / 1024;
                                    pic.PictureFormat.Crop.PictureOffsetY = (pic.PictureFormat.Crop.ShapeHeight - pic.PictureFormat.Crop.PictureHeight) / 1024;
                                    label1.Text = Math.Round(pic.Width / pic.Height, 2) + " : 1";
                                }
                                else
                                {
                                    float h = float.Parse(arr[0]);
                                    float v = float.Parse(arr[1]);

                                    if (h >= pic.Width / pic.Height * v)
                                    {
                                        pic.PictureFormat.Crop.PictureOffsetY = -(pic.Height - v / h * pic.Width) / 2;
                                        pic.PictureFormat.Crop.ShapeHeight = v / h * pic.Width;
                                    }
                                    else
                                    {
                                        pic.PictureFormat.Crop.PictureOffsetX = -(pic.Width - h / v * pic.Height) / 2;
                                        pic.PictureFormat.Crop.ShapeWidth = h / v * pic.Height;
                                    }
                                    label1.Text = h + " : " + v;
                                }
                                pic.Left = lm - pic.Width / 2;
                                pic.Top = tm - pic.Height / 2;
                            }
                        }
                    }
                }

                if (comboBox1.SelectedItem as string == "填充色单值统一(回车)")
                {
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                    {
                        label1.Text = "色型,递进值";
                    }
                    else
                    {
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        if (sel.HasChildShapeRange)
                        {
                            range = sel.ChildShapeRange;
                        }
                        int count = range.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            PowerPoint.Shape shape = range[i];
                            if (shape.Type == Office.MsoShapeType.msoGroup)
                            {
                                int count2 = shape.GroupItems.Count;
                                for (int j = 1; j <= count2; j++)
                                {
                                    PowerPoint.Shape gshape = shape.GroupItems[j];
                                    if (gshape.Shadow.Visible == Office.MsoTriState.msoTrue)
                                    {
                                        float tr = gshape.Shadow.Transparency;
                                        int rgb0 = gshape.Shadow.ForeColor.RGB;
                                        int nrgb = Nrgb3(rgb0);
                                        gshape.Shadow.ForeColor.RGB = nrgb;
                                        gshape.Shadow.Transparency = tr;
                                    }
                                    if (gshape.Fill.Type == Office.MsoFillType.msoFillGradient)
                                    {
                                        int count3 = gshape.Fill.GradientStops.Count;
                                        for (int k = 1; k <= count3; k++)
                                        {
                                            int rgb0 = gshape.Fill.GradientStops[k].Color.RGB;
                                            float tr = gshape.Fill.GradientStops[k].Transparency;
                                            int nrgb = Nrgb3(rgb0);
                                            gshape.Fill.GradientStops[k].Color.RGB = nrgb;
                                            gshape.Fill.GradientStops[k].Transparency = tr;
                                        }
                                    }
                                    else if (gshape.Fill.Type == Office.MsoFillType.msoFillSolid)
                                    {
                                        float tr = gshape.Fill.Transparency;
                                        int rgb0 = gshape.Fill.ForeColor.RGB;
                                        int nrgb = Nrgb3(rgb0);
                                        gshape.Fill.ForeColor.RGB = nrgb;
                                        gshape.Fill.Transparency = tr;
                                    }
                                    else
                                    {
                                        label1.Text = "仅支持纯色和渐变";
                                    }
                                }
                            }
                            else
                            {
                                if (shape.Shadow.Visible == Office.MsoTriState.msoTrue)
                                {
                                    float tr = shape.Shadow.Transparency;
                                    int rgb0 = shape.Shadow.ForeColor.RGB;
                                    int nrgb = Nrgb3(rgb0);
                                    shape.Shadow.ForeColor.RGB = nrgb;
                                    shape.Shadow.Transparency = tr;
                                }
                                if (shape.Fill.Type == Office.MsoFillType.msoFillGradient)
                                {
                                    int count4 = shape.Fill.GradientStops.Count;
                                    for (int m = 1; m <= count4; m++)
                                    {
                                        int rgb0 = shape.Fill.GradientStops[m].Color.RGB;
                                        float tr = shape.Fill.GradientStops[m].Transparency;
                                        int nrgb = Nrgb3(rgb0);
                                        shape.Fill.GradientStops[m].Color.RGB = nrgb;
                                        shape.Fill.GradientStops[m].Transparency = tr;
                                    }
                                }
                                else if (shape.Fill.Type == Office.MsoFillType.msoFillSolid)
                                {
                                    float tr = shape.Fill.Transparency;
                                    int rgb0 = shape.Fill.ForeColor.RGB;
                                    int nrgb = Nrgb3(rgb0);
                                    shape.Fill.ForeColor.RGB = nrgb;
                                    shape.Fill.Transparency = tr;
                                }
                                else
                                {
                                    label1.Text = "仅支持纯色和渐变";
                                }
                            }
                        }
                    }
                }

                if (comboBox1.SelectedItem as string == "线条色单值统一(回车)")
                {
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                    {
                        label1.Text = "色型,递进值";
                    }
                    else
                    {
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        if (sel.HasChildShapeRange)
                        {
                            range = sel.ChildShapeRange;
                        }
                        int count = range.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            PowerPoint.Shape shape = range[i];
                            if (shape.Type == Office.MsoShapeType.msoGroup)
                            {
                                int count2 = shape.GroupItems.Count;
                                for (int j = 1; j <= count2; j++)
                                {
                                    PowerPoint.Shape gshape = shape.GroupItems[j];
                                    if (gshape.Line.Visible == Office.MsoTriState.msoTrue)
                                    {
                                        float tr = gshape.Line.Transparency;
                                        int rgb0 = gshape.Line.ForeColor.RGB;
                                        int nrgb = Nrgb3(rgb0);
                                        gshape.Line.ForeColor.RGB = nrgb;
                                        gshape.Line.Transparency = tr;
                                    }
                                    else
                                    {
                                        label1.Text = "先添加线条色";
                                    }
                                }
                            }
                            else
                            {
                                if (shape.Line.Visible == Office.MsoTriState.msoTrue)
                                {
                                    float tr = shape.Line.Transparency;
                                    int rgb0 = shape.Line.ForeColor.RGB;
                                    int nrgb = Nrgb3(rgb0);
                                    shape.Line.ForeColor.RGB = nrgb;
                                    shape.Line.Transparency = tr;
                                }
                                else
                                {
                                    label1.Text = "先添加线条色";
                                }
                            }
                        }
                    }
                }

                if (comboBox1.SelectedItem as string == "阴影色单值统一(回车)")
                {
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                    {
                        label1.Text = "色型,递进值";
                    }
                    else
                    {
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        if (sel.HasChildShapeRange)
                        {
                            range = sel.ChildShapeRange;
                        }
                        int count = range.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            PowerPoint.Shape shape = range[i];
                            if (shape.Type == Office.MsoShapeType.msoGroup)
                            {
                                int count2 = shape.GroupItems.Count;
                                for (int j = 1; j <= count2; j++)
                                {
                                    PowerPoint.Shape gshape = shape.GroupItems[j];
                                    if (gshape.Shadow.Visible == Office.MsoTriState.msoTrue)
                                    {
                                        float tr = gshape.Shadow.Transparency;
                                        int rgb0 = gshape.Shadow.ForeColor.RGB;
                                        int nrgb = Nrgb3(rgb0);
                                        gshape.Shadow.ForeColor.RGB = nrgb;
                                        gshape.Shadow.Transparency = tr;
                                    }
                                    else
                                    {
                                        label1.Text = "先设置阴影";
                                    }
                                }
                            }
                            else
                            {
                                if (shape.Shadow.Visible == Office.MsoTriState.msoTrue)
                                {
                                    float tr = shape.Shadow.Transparency;
                                    int rgb0 = shape.Shadow.ForeColor.RGB;
                                    int nrgb = Nrgb3(rgb0);
                                    shape.Shadow.ForeColor.RGB = nrgb;
                                    shape.Shadow.Transparency = tr;
                                }
                                else
                                {
                                    label1.Text = "先设置阴影";
                                }
                            }
                        }
                    }
                }

                if (comboBox1.SelectedItem as string == "-------------------------")
                {
                    int picmix=int.Parse(textBox1.Text.Trim());
                    if (picmix == 0)
                    {
                        Properties.Settings.Default.PicMix = 0;
                        Properties.Settings.Default.Save();
                        MessageBox.Show("已恢复默认");
                    }
                    else if (picmix == 1)
                    {
                        Properties.Settings.Default.PicMix = 1;
                        Properties.Settings.Default.Save();
                        MessageBox.Show("已修改成功");
                    }
                }

                if (comboBox1.SelectedItem as string == "批量改名(回车)")
                {
                    string name = textBox1.Text.Trim();
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        if (sel.HasChildShapeRange)
                        {
                            range = sel.ChildShapeRange;
                        }
                        if (name.Contains(char.Parse("_")) || name.Contains(char.Parse("-")) || name == "")
                        {
                            for (int i = 1; i <= range.Count; i++)
                            {
                                range[i].Name = name + i;
                            }
                        }
                        else
                        {
                            for (int i = 1; i <= range.Count; i++)
                            {
                                range[i].Name = name;
                            }
                        }
                    }
                }

                if (comboBox1.SelectedItem as string == "批量加字(回车)")
                {
                    if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        PowerPoint.ShapeRange range = sel.ShapeRange;
                        if (sel.HasChildShapeRange)
                        {
                            range = sel.ChildShapeRange;
                        }
                        string[] arr = textBox1.Text.Trim().Split(char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
                        for (int i = 1; i <= range.Count; i++)
                        {
                            PowerPoint.Shape shape = range[i];
                            if (shape.Type == Office.MsoShapeType.msoGroup)
                            {
                                int gn = 0;
                                foreach (PowerPoint.Shape nshape in shape.GroupItems)
                                {
                                    if (nshape.TextEffect.Text != "")
                                    {
                                        nshape.TextEffect.Text = arr[(i - 1) % arr.Count()];
                                    }
                                    else
                                    {
                                        gn += 1;
                                    }
                                }
                                if (gn == shape.GroupItems.Count)
                                {
                                    shape.GroupItems[shape.GroupItems.Count].TextEffect.Text = arr[(i - 1) % arr.Count()];
                                }
                            }
                            else
                            {
                                shape.TextEffect.Text = arr[(i - 1) % arr.Count()];
                            }
                        }
                    }
                }

                if (comboBox1.SelectedItem as string == "文本字号递进(回车)")
                {
                    string[] arr = textBox1.Text.Trim().Split(char.Parse(":"), char.Parse("："), char.Parse(" "), char.Parse(","), char.Parse("，")).ToArray();
                    if (arr.Count() == 1)
                    {
                        if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                        {
                            float ts = float.Parse(arr[0]);
                            PowerPoint.ShapeRange range = sel.ShapeRange;
                            if (sel.HasChildShapeRange)
                            {
                                range = sel.ChildShapeRange;
                            }
                            if (range.Count == 1)
                            {
                                if (range[1].Type == Office.MsoShapeType.msoGroup)
                                {
                                    List<PowerPoint.Shape> gshapes = new List<PowerPoint.Shape>();
                                    foreach (PowerPoint.Shape gshape in range[1].GroupItems)
                                    {
                                        if (gshape.HasTextFrame == Office.MsoTriState.msoTrue && gshape.TextFrame.HasText== Office.MsoTriState.msoTrue && gshape.TextFrame.TextRange.Text != "")
                                        {
                                            gshapes.Add(gshape);
                                        }
                                    }
                                    if (gshapes.Count > 0)
                                    {
                                        for (int i = 0; i < gshapes.Count; i++)
                                        {
                                            gshapes[i].TextFrame.TextRange.Font.Size = ts;
                                        }
                                        label1.Text = "字号: " + ts;
                                    }
                                    else
                                    {
                                        label1.Text = "请选中文本";
                                    }
                                
                                }
                                else
                                {
                                    if (range[1].HasTextFrame == Office.MsoTriState.msoTrue && range[1].TextFrame.HasText == Office.MsoTriState.msoTrue && range[1].TextFrame.TextRange.Text != "")
                                    {
                                        PowerPoint.TextRange tr = range[1].TextFrame.TextRange;
                                        int nc = tr.Text.Count();
                                        for (int i = 1; i <= nc; i++)
                                        {
                                            tr.Characters(i).Font.Size = ts;
                                        }
                                        label1.Text = "字号: " + ts;
                                    }
                                    else
                                    {
                                        label1.Text = "请选中文本";
                                    }
                                }
                            }
                            else
                            {
                                List<PowerPoint.Shape> shapes = new List<PowerPoint.Shape>();
                                foreach (PowerPoint.Shape shape in range)
                                {
                                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue && shape.TextFrame.TextRange.Text != "")
                                    {
                                        shapes.Add(shape);
                                    }
                                }
                                if (shapes.Count > 0)
                                {
                                    for (int i = 0; i < shapes.Count; i++)
                                    {
                                        shapes[i].TextFrame.TextRange.Font.Size = ts;
                                    }
                                    label1.Text = "字号: " + ts;
                                }
                                else
                                {
                                    label1.Text = "请选中文本";
                                }
                            }
                        }
                        else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                        {  
                            float ts = float.Parse(arr[0]);
                            PowerPoint.TextRange tr = sel.TextRange;
                            int nc = tr.Text.Count();
                            for (int i = 1; i <= nc; i++)
                            {
                                tr.Characters(i).Font.Size = ts;
                            }
                            label1.Text = "字号: " + ts;
                        }
                        else
                        {
                            label1.Text = "请选中文本";
                        }
                    }
                    else if (arr.Count() == 2)
                    {
                        if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                        {
                            float min = float.Parse(arr[0]); float max = float.Parse(arr[1]);
                            PowerPoint.ShapeRange range = sel.ShapeRange;
                            if (sel.HasChildShapeRange)
                            {
                                range = sel.ChildShapeRange;
                            }
                            if (range.Count == 1)
                            {
                                if (range[1].Type == Office.MsoShapeType.msoGroup)
                                {
                                    List<PowerPoint.Shape> gshapes = new List<PowerPoint.Shape>();
                                    foreach (PowerPoint.Shape gshape in range[1].GroupItems)
                                    {
                                        if (gshape.HasTextFrame == Office.MsoTriState.msoTrue && gshape.TextFrame.HasText== Office.MsoTriState.msoTrue && gshape.TextFrame.TextRange.Text != "")
                                        {
                                            gshapes.Add(gshape);
                                        }
                                    }
                                    if (gshapes.Count > 0)
                                    {
                                        float n = (max - min) / (float)(gshapes.Count - 1);
                                        for (int i = 0; i < gshapes.Count; i++)
                                        {
                                            gshapes[i].TextFrame.TextRange.Font.Size = min + n * i;
                                        }
                                        label1.Text = "字号: " + min + " 到 " + max;
                                    }
                                    else
                                    {
                                        label1.Text = "请选中文本";
                                    }
                                
                                }
                                else
                                {
                                    if (range[1].HasTextFrame == Office.MsoTriState.msoTrue && range[1].TextFrame.HasText == Office.MsoTriState.msoTrue && range[1].TextFrame.TextRange.Text != "")
                                    {
                                        PowerPoint.TextRange tr = range[1].TextFrame.TextRange;
                                        int nc = tr.Text.Count();
                                        float n = (max - min) / (float)(nc - 1);
                                        for (int i = 1; i <= nc; i++)
                                        {
                                            tr.Characters(i).Font.Size = min + n * (i - 1);
                                        }
                                        label1.Text = "字号: " + min + " 到 " + max;
                                    }
                                    else
                                    {
                                        label1.Text = "请选中文本";
                                    }
                                }
                            }
                            else
                            {
                                List<PowerPoint.Shape> shapes = new List<PowerPoint.Shape>();
                                foreach (PowerPoint.Shape shape in range)
                                {
                                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue && shape.TextFrame.TextRange.Text != "")
                                    {
                                        shapes.Add(shape);
                                    }
                                }
                                if (shapes.Count > 0)
                                {
                                    float n = (max - min) / (float)(shapes.Count - 1);
                                    for (int i = 0; i < shapes.Count; i++)
                                    {
                                        shapes[i].TextFrame.TextRange.Font.Size = min + n * i;
                                    }
                                    label1.Text = "字号: " + min + " 到 " + max;
                                }
                                else
                                {
                                    label1.Text = "请选中文本";
                                }
                            }
                        }
                        else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                        {
                            float min = float.Parse(arr[0]); float max = float.Parse(arr[1]);
                            PowerPoint.TextRange tr = sel.TextRange;
                            int nc = tr.Text.Count();
                            float n = (max - min) / (float)(nc - 1);
                            for (int i = 1; i <= nc; i++)
                            {
                                tr.Characters(i).Font.Size = min + n * (i - 1);
                            }
                            label1.Text = "字号: " + min + " 到 " + max;
                        }
                        else
                        {
                            label1.Text = "请选中文本";
                        }
                    }
                    else
                    {
                        label1.Text = "仅需2个值";
                    }
                }
                    
                e.Handled = true;
            }
  
        }

    }
}
