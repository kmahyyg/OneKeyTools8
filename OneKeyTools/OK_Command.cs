using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Speech.Recognition;
using System.Globalization;
using NAudio.Wave;
using NAudio;
using Microsoft.Win32;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace OneKeyTools
{
    public partial class OK_Command : Form
    {
        public OK_Command()
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

        private SpeechRecognitionEngine SRE = null;
        private DictationGrammar gm = null;
        private WaveFileReader wfr = null;
        private WaveOut waveout = null;

        private void OKRecognition2_Load(object sender, EventArgs e)
        {
            SRE = new SpeechRecognitionEngine(SpeechRecognitionEngine.InstalledRecognizers()[0]);
            gm = new DictationGrammar();
            SRE.LoadGrammar(gm);
            SRE.SetInputToDefaultAudioDevice();
            SRE.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(recognizer_SpeechRecognized);
            if (Properties.Settings.Default.CommandSound == true)
            {
                静音ToolStripMenuItem.Text = "静音";
            }
            else
            {
                静音ToolStripMenuItem.Text = "取消静音";
            }
        }

        private void TipSound()
        {
            if (Properties.Settings.Default.CommandSound == true)
            {
                if (wfr != null)
                {
                    wfr.Dispose();
                }
                if (waveout != null)
                {
                    waveout.Dispose();
                }
                wfr = new WaveFileReader(Properties.Resources.Success_Rec);
                waveout = new WaveOut();
                waveout.Init(wfr);
                waveout.Play();
            }
        }

        private void recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            OKCommand.Text = e.Result.Text;
            switch (e.Result.Text)
            {
                case "停止识别":
                    TipSound();
                    OKCommand_Click(null,null);
                    break;
                case "关于":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button1_Click(null, null);
                    break;
                case "锐普论坛":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button81_Click(null, null);
                    break;
                case "演界网":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button83_Click(null, null);
                    break;
                case "Office中国":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button2_Click(null, null);
                    break;
                case "教程合辑":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button82_Click(null, null);
                    break;
                case "长图教程":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button84_Click(null, null);
                    break;
                case "安装位置":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button193_Click(null, null);
                    break;
                case "设置":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button127_Click(null, null);
                    break;
                case "官网":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button79_Click(null, null);
                    break;
                case "关注":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button128_Click(null, null);
                    break;
                case "隐藏副卡":
                    MessageBox.Show("使用OK命令时，不能语音隐藏或显示副卡");
                    break;
                case "显示副卡":
                    MessageBox.Show("使用OK命令时，不能语音隐藏或显示副卡");
                    break;
                case "隐藏主卡":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button247_Click(null, null);
                    break;
                case "显示主卡":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button247_Click(null, null);
                    break;
                case "插入形状":
                    TipSound();
                    Globals.Ribbons.Ribbon1.splitButton3_Click(null, null);
                    break;
                case "全屏矩形":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button8_Click(null, null);
                    break;
                case "插入圆形":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button9_Click(null, null);
                    break;
                case "EMF":
                    TipSound();
                    Globals.Ribbons.Ribbon1.splitButton4_Click(null, null);
                    break;
                case "EMF导入":
                    TipSound();
                    Globals.Ribbons.Ribbon1.splitButton4_Click(null, null);
                    break;
                case "导入后独立":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button18_Click(null, null);
                    break;
                case "导入后组合":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button19_Click(null, null);
                    break;
                case "去形状":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button12_Click(null, null);
                    break;
                case "去占位符":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button14_Click(null, null);
                    break;
                case "去同位":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button232_Click(null, null);
                    break;
                case "去版式":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button282_Click(null, null);
                    break;
                case "去形状填充":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button15_Click(null, null);
                    break;
                case "去形状边框":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button16_Click(null, null);
                    break;
                case "去形状阴影":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button17_Click(null, null);
                    break;
                case "去文字":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button13_Click(null, null);
                    break;
                case "去文本边框":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button174_Click(null, null);
                    break;
                case "去表格文本":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button278_Click(null, null);
                    break;
                case "去形状字体":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button283_Click(null, null);
                    break;
                case "去图片":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button155_Click(null, null);
                    break;
                case "去音频":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button175_Click(null, null);
                    break;
                case "去视频":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button211_Click(null, null);
                    break;
                case "去图表":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button275_Click(null, null);
                    break;
                case "去表格":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button276_Click(null, null);
                    break;
                case "锁定纵横比":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button137_Click(null, null);
                    break;
                case "取消纵横比":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button139_Click(null, null);
                    break;
                case "去超链接":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button116_Click(null, null);
                    break;
                case "单页组合":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button241_Click(null, null);
                    break;
                case "去组合":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button235_Click(null, null);
                    break;
                case "相同大小":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button23_Click(null, null);
                    break;
                case "从小到大":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button21_Click(null, null);
                    break;
                case "从大到小":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button22_Click(null, null);
                    break;
                case "统一线宽":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button200_Click(null, null);
                    break;
                case "从窄到宽":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button201_Click(null, null);
                    break;
                case "从宽到窄":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button199_Click(null, null);
                    break;
                case "随机线宽":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button197_Click(null, null);
                    break;
                case "随机大小":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button115_Click(null, null);
                    break;
                case "随机位置":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button138_Click(null, null);
                    break;
                case "横纵方向":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button221_Click(null, null);
                    break;
                case "横方向":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button222_Click(null, null);
                    break;
                case "纵方向":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button222_Click(null, null);
                    break;
                case "对齐递进":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button20_Click(null, null);
                    break;
                case "对齐增强":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button194_Click(null, null);
                    break;
                case "经典对齐":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button202_Click(null, null);
                    break;
                case "全屏大小":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button168_Click(null, null);
                    break;
                case "左顶对齐":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button161_Click(null, null);
                    break;
                case "居中对齐":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button162_Click(null, null);
                    break;
                case "旋转递进":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button24_Click(null, null);
                    break;
                case "随机旋转":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button192_Click(null, null);
                    break;
                case "旋转增强":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button28_Click(null, null);
                    break;
                case "控点工具":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button30_Click(null, null);
                    break;
                case "矩式复制":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button29_Click(null, null);
                    break;
                case "环式复制":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button186_Click(null, null);
                    break;
                case "路径复制":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button190_Click(null, null);
                    break;
                case "尺寸复制":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button212_Click(null, null);
                    break;
                case "文本统一":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button85_Click(null, null);
                    break;
                case "按段拆分":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button31_Click(null, null);
                    break;
                case "合并段落":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button32_Click(null, null);
                    break;
                case "拆分增强":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button184_Click(null, null);
                    break;
                case "拆为单字":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button33_Click(null, null);
                    break;
                case "单字合并":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button34_Click(null, null);
                    break;
                case "原位复制":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button169_Click(null, null);
                    break;
                case "三角分形":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button132_Click(null, null);
                    break;
                case "正方形分形":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button133_Click(null, null);
                    break;
                case "梯形分形":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button134_Click(null, null);
                    break;
                case "超级折线":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button245_Click(null, null);
                    break;
                case "辐射连线":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button213_Click(null, null);
                    break;
                case "就近连线":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button215_Click(null, null);
                    break;
                case "组合连线":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button214_Click(null, null);
                    break;
                case "顶点计数":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button131_Click(null, null);
                    break;
                case "顶点横平":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button129_Click(null, null);
                    break;
                case "顶点竖直":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button130_Click(null, null);
                    break;
                case "抻直弓形":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button144_Click(null, null);
                    break;
                case "二等分点":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button158_Click(null, null);
                    break;
                case "三等分点":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button159_Click(null, null);
                    break;
                case "去除等分点":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button160_Click(null, null);
                    break; 
                case "删除备注":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button165_Click(null, null);
                    break;
                case "备注合并":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button166_Click(null, null);
                    break;
                case "备注导入":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button167_Click(null, null);
                    break;
                case "补加备注":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button173_Click(null, null);
                    break;
                case "纯色统一":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button93_Click(null, null);
                    break;
                case "HSL补色":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button39_Click(null, null);
                    break;
                case "H补色":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button40_Click(null, null);
                    break;
                case "S补色":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button41_Click(null, null);
                    break;
                case "L补色":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button42_Click(null, null);
                    break;
                case "RGB补色":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button43_Click(null, null);
                    break;
                case "R补色":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button44_Click(null, null);
                    break;
                case "G补色":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button45_Click(null, null);
                    break;
                case "B补色":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button46_Click(null, null);
                    break;
                case "随机纯色":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button191_Click(null, null);
                    break;
                case "随机填充透明":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button198_Click(null, null);
                    break;
                case "渐变转纯色":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button35_Click(null, null);
                    break;
                case "纯色转渐变":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button36_Click(null, null);
                    break;
                case "光圈同色":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button98_Click(null, null);
                    break;
                case "光圈虚化":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button140_Click(null, null);
                    break;
                case "光圈锐化":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button260_Click(null, null);
                    break;
                case "光圈等分":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button189_Click(null, null);
                    break;
                case "光圈逆序":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button78_Click(null, null);
                    break;
                case "去除光圈":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button145_Click(null, null);
                    break;
                case "RGB色值":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button37_Click(null, null);
                    break;
                case "HSL色值":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button38_Click(null, null);
                    break;
                case "16进制色值":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button236_Click(null, null);
                    break;
                case "取色器":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button80_Click(null, null);
                    break;
                case "填线到阴影":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button122_Click(null, null);
                    break;
                case "文字到阴影":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button243_Click(null, null);
                    break;
                case "填线到文字":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button10_Click(null, null);
                    break;
                case "阴影到文字":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button244_Click(null, null);
                    break;
                case "线条到填充":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button123_Click(null, null);
                    break;
                case "阴影到填充":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button124_Click(null, null);
                    break;
                case "文字到填充":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button11_Click(null, null);
                    break;
                case "填充到线条":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button121_Click(null, null);
                    break;
                case "阴影到线条":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button125_Click(null, null);
                    break;
                case "文字到线条":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button242_Click(null, null);
                    break;
                case "OK神框":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button90_Click(null, null);
                    break;
                case "神框":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button90_Click(null, null);
                    break;
                case "三维复制":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button57_Click(null, null);
                    break;
                case "三维全刷":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button54_Click(null, null);
                    break;
                case "三维旋转刷":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button55_Click(null, null);
                    break;
                case "三维格式刷":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button56_Click(null, null);
                    break;
                case "检测旋转类型":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button136_Click(null, null);
                    break;
                case "三维演示":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button142_Click(null, null);
                    break;
                case "去除形状三维":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button47_Click(null, null);
                    break;
                case "一键球体":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button48_Click(null, null);
                    break;
                case "一键立方体":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button49_Click(null, null);
                    break;
                case "一键水晶体":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button156_Click(null, null);
                    break;
                case "沙漪立方拼":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button50_Click(null, null);
                    break;
                case "水平补位":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button51_Click(null, null);
                    break;
                case "垂直补位":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button52_Click(null, null);
                    break;
                case "批量补位":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button210_Click(null, null);
                    break;
                case "添加透视":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button53_Click(null, null);
                    break;
                case "图填充到形状":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button68_Click(null, null);
                    break;
                case "图片反相":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button105_Click(null, null);
                    break;
                case "三通道分离":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button117_Click(null, null);
                    break;
                case "三通道合并":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button114_Click(null, null);
                    break;
                case "色调置换":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button195_Click(null, null);
                    break;
                case "饱和度置换":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button227_Click(null, null);
                    break;
                case "亮度置换":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button228_Click(null, null);
                    break;
                case "红通道置换":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button229_Click(null, null);
                    break;
                case "绿通道置换":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button230_Click(null, null);
                    break;
                case "蓝通道置换":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button231_Click(null, null);
                    break;
                case "变暗":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button96_Click(null, null);
                    break;
                case "正片叠底":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button4_Click(null, null);
                    break;
                case "颜色加深":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button5_Click(null, null);
                    break;
                case "线性加深":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button106_Click(null, null);
                    break;
                case "变亮":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button7_Click(null, null);
                    break;
                case "滤色":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button97_Click(null, null);
                    break;
                case "颜色减淡":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button6_Click(null, null);
                    break;
                case "线性减淡":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button107_Click(null, null);
                    break;
                case "叠加":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button108_Click(null, null);
                    break;
                case "柔光":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button109_Click(null, null);
                    break;
                case "强光":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button110_Click(null, null);
                    break;
                case "亮光":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button119_Click(null, null);
                    break;
                case "图片虚化":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button63_Click(null, null);
                    break;
                case "三维折图":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button69_Click(null, null);
                    break;
                case "马赛克":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button3_Click(null, null);
                    break;
                case "形状裁图":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button88_Click(null, null);
                    break;
                case "图片极坐标":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button141_Click(null, null);
                    break;
                case "剪影辅助":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button187_Click(null, null);
                    break;
                case "图片画中画":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button171_Click(null, null);
                    break;
                case "图片弧化":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button205_Click(null, null);
                    break;
                case "图片倾斜":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button216_Click(null, null);
                    break;
                case "形状模糊":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button64_Click(null, null);
                    break;
                case "形状吸附路径":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button102_Click(null, null);
                    break;
                case "形状取像素":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button104_Click(null, null);
                    break;
                case "水滴质感":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button66_Click(null, null);
                    break;
                case "军事迷彩":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button65_Click(null, null);
                    break;
                case "页面导图":
                    TipSound();
                    Globals.Ribbons.Ribbon1.splitButton2_Click(null, null);
                    break;
                case "快捷拼图":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button171_Click(null, null);
                    break;
                case "自由拼图":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button170_Click(null, null);
                    break;
                case "微信封面":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button163_Click(null, null);
                    break;
                case "导图设置":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button126_Click(null, null);
                    break;
                case "新建库":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button206_Click(null, null);
                    break;
                case "从库删除":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button208_Click(null, null);
                    break;
                case "设置库目录":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button219_Click(null, null);
                    break;
                case "恢复库目录":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button220_Click(null, null);
                    break;
                case "加载库":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button207_Click(null, null);
                    break;
                case "添加到库":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button209_Click(null, null);
                    break;
                case "一键转图":
                    TipSound();
                    Globals.Ribbons.Ribbon1.splitButton7_Click(null, null);
                    break;
                case "原位转PNG":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button58_Click(null, null);
                    break;
                case "原位转JPG":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button59_Click(null, null);
                    break;
                case "批量导PNG":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button61_Click(null, null);
                    break;
                case "批量导JPG":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button185_Click(null, null);
                    break;
                case "GIF工具":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button60_Click(null, null);
                    break;
                case "裁图辅助":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button67_Click(null, null);
                    break;
                case "按尺寸":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button95_Click(null, null);
                    break;
                case "按类型":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button147_Click(null, null);
                    break;
                case "按自选图形":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button238_Click(null, null);
                    break;
                case "按填充类型":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button101_Click(null, null);
                    break;
                case "按渐变色":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button150_Click(null, null);
                    break;
                case "按光圈数":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button148_Click(null, null);
                    break;
                case "按填充纯色":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button149_Click(null, null);
                    break;
                case "按填充透明度":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button146_Click(null, null);
                    break;
                case "按线条纯色":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button152_Click(null, null);
                    break;
                case "按线条透明度":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button153_Click(null, null);
                    break;
                case "按线条粗细":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button151_Click(null, null);
                    break;
                case "反选对象":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button224_Click(null, null);
                    break;
                case "隔选对象":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button225_Click(null, null);
                    break;
                case "随机选择":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button226_Click(null, null);
                    break;
                case "所选统计":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button143_Click(null, null);
                    break;
                case "隐藏所选":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button103_Click(null, null);
                    break;
                case "全部显示":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button234_Click(null, null);
                    break;
                case "二等分线":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button91_Click(null, null);
                    break;
                case "左上黄金线":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button70_Click(null, null);
                    break;
                case "右下黄金线":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button71_Click(null, null);
                    break;
                case "垂直三分线":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button86_Click(null, null);
                    break;
                case "水平三分线":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button87_Click(null, null);
                    break;
                case "去分割线":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button72_Click(null, null);
                    break;
                case "图换形":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button203_Click(null, null);
                    break;
                case "图换形循环":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button204_Click(null, null);
                    break;
                case "图换图":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button217_Click(null, null);
                    break;
                case "同名换图":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button261_Click(null, null);
                    break;
                case "形换形循环":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button218_Click(null, null);
                    break;
                case "保留格式":
                    TipSound();
                    Globals.Ribbons.Ribbon1.checkBox1_Click(null, null);
                    break;
                case "选图分页1":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button89_Click(null, null);
                    break;
                case "选图分页2":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button196_Click(null, null);
                    break;
                case "尺寸多页统一":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button100_Click(null, null);
                    break;
                case "位置多页统一":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button99_Click(null, null);
                    break;
                case "合并图片":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button233_Click(null, null);
                    break;
                case "合并文本框":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button237_Click(null, null);
                    break;
                case "合并形状":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button239_Click(null, null);
                    break;
                case "合并图表":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button273_Click(null, null);
                    break;
                case "合并表格":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button274_Click(null, null);
                    break;
                case "合并所有":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button240_Click(null, null);
                    break;
                case "吸附对齐":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button135_Click(null, null);
                    break;
                case "超级缩放":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button157_Click(null, null);
                    break;
                case "演示白板":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button181_Click(null, null);
                    break;
                case "沐心放映":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button27_Click(null, null);
                    break;
                case "计算器":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button77_Click(null, null);
                    break;
                case "记事本":
                    TipSound();
                    Notepad();
                    break;
                case "数字时钟":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button25_Click(null, null);
                    break;
                case "定时器":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button26_Click(null, null);
                    break;
                case "倒计时":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button154_Click(null, null);
                    break;
                case "形线相连":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button176_Click(null, null);
                    break;
                case "顶点微调":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button177_Click(null, null);
                    break;
                case "顶点重合":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button172_Click(null, null);
                    break;
                case "激活三维":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button188_Click(null, null);
                    break;
                case "形状裁图2":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button182_Click(null, null);
                    break;
                case "小影正片叠底":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button92_Click(null, null);
                    break;
                case "小二微立体":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button62_Click(null, null);
                    break;
                case "小天逆时针形状":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button73_Click(null, null);
                    break;
                case "小天逆时针圆环":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button74_Click(null, null);
                    break;
                case "书馨形状长阴影":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button75_Click(null, null);
                    break;
                case "书馨文字长阴影":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button76_Click(null, null);
                    break;
                case "文字透明1":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button94_Click(null, null);
                    break;
                case "文字透明2":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button118_Click(null, null);
                    break;
                case "字符加空":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button280_Click(null, null);
                    break;
                case "字符减空":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button281_Click(null, null);
                    break;
                case "黑洞微博":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button120_Click(null, null);
                    break;
                case "图形逐帧":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button179_Click(null, null);
                    break;
                case "逐帧辅助":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button183_Click(null, null);
                    break;
                case "口袋官网":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button180_Click(null, null);
                    break;
                case "插入表格":
                    TipSound();
                    Globals.Ribbons.Ribbon1.splitButton8_Click(null, null);
                    break;
                case "表格设置":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button249_Click(null, null);
                    break;
                case "表格上色":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button248_Click(null, null);
                    break;
                case "列求和":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button252_Click(null, null);
                    break;
                case "行求和":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button253_Click(null, null);
                    break;
                case "找最大值":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button255_Click(null, null);
                    break;
                case "找最小值":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button256_Click(null, null);
                    break;
                case "向下取整":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button257_Click(null, null);
                    break;
                case "向上取整":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button258_Click(null, null);
                    break;
                case "取绝对值":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button259_Click(null, null);
                    break;
                case "计算增强":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button254_Click(null, null);
                    break;
                case "字转表格":
                    TipSound();
                    Globals.Ribbons.Ribbon1.splitButton9_Click(null, null);
                    break;
                case "按段转换":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button250_Click(null, null);
                    break;
                case "提取文本":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button251_Click(null, null);
                    break;
                case "表格边框":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button262_Click(null, null);
                    break;
                case "保存图表样式":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button178_Click(null, null);
                    break;
                case "套用图表样式":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button263_Click(null, null);
                    break;
                case "设置图表目录":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button264_Click(null, null);
                    break;
                case "来自电子表格":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button265_Click(null, null);
                    break;
                case "来自文档":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button279_Click(null, null);
                    break;
                case "来自剪贴板":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button266_Click(null, null);
                    break;
                case "取消图表链接":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button267_Click(null, null);
                    break;
                case "图表转表格":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button267_Click(null, null);
                    break;
                case "表格转图表":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button267_Click(null, null);
                    break;
                case "图表导出为PNG":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button267_Click(null, null);
                    break;
                case "图表导出为JPG":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button267_Click(null, null);
                    break;
                case "导出到电子表格":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button267_Click(null, null);
                    break;
                case "导出到文档":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button267_Click(null, null);
                    break;
                case "粘贴到备注":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button111_Click(null, null);
                    break;
                case "粘贴到占位符":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button112_Click(null, null);
                    break;
                case "粘贴到母版":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button113_Click(null, null);
                    break;
                case "朗读工具":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button285_Click(null, null);
                    break;
                case "音频拆合":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button284_Click(null, null);
                    break;
                case "音频混合":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button288_Click(null, null);
                    break;
                case "音频转换":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button289_Click(null, null);
                    break;
                case "录音工具":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button286_Click(null, null);
                    break;
                case "合并时保留格式":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button290_Click(null, null);
                    break;
                case "合并时不保留格式":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button291_Click(null, null);
                    break;
                case "合并首页":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button292_Click(null, null);
                    break;
                case "截取页面":
                    TipSound();
                    Globals.Ribbons.Ribbon1.button293_Click(null, null);
                    break;
                case "水平居中":
                    TipSound();
                    HAlign();
                    break;
                case "垂直居中":
                    TipSound();
                    VAlign();
                    break;
                case "全选":
                    TipSound();
                    selectallshapes();
                    break;
                case "全部选中":
                    TipSound();
                    selectallshapes();
                    break;
                case "全选页面":
                    TipSound();
                    selectallslides();
                    break;
                case "复制页面":
                    TipSound();
                    CopySlide();
                    break;
                case "水平贴边":
                    TipSound();
                    Align_HTiebian();
                    break;
                case "垂直贴边":
                    TipSound();
                    Align_VTiebian();
                    break;
            }
        }

        private void 关闭ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (SRE!= null)
            {
                SRE.Dispose();
                SRE = null;
            }
            OK_Command.ActiveForm.Close();
            Globals.Ribbons.Ribbon1.button287.Enabled = true;
        }

        private void OKCommand_Click(object sender, EventArgs e)
        {
            if (OKCommand.Text == "单击此处" || OKCommand.Text == "已停止")
            {
                SRE.RecognizeAsync(RecognizeMode.Multiple);
                OKCommand.Text = "请说话";
            }
            else
            {
                if (SRE != null)
                {
                    SRE.RecognizeAsyncStop();
                    SRE.RecognizeAsyncCancel();
                    OKCommand.Text = "已停止";
                }
            }
        }

        private void 说明ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("OK命令基于Windows系统内置的语音识别引擎，目前支持OK在功能区上的所有功能（除图形库、隐藏副卡、OK命令外）的语音调用。若要提高准确度，请到官网（http://oktools.xyz）查看语音配置文件导入方法");
        }

        PowerPoint.Application app = Globals.ThisAddIn.Application;
        public void HAlign()
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
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
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
            else
            {
                MessageBox.Show("请选中形状");
            }
        }

        public void VAlign()
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
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
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
            else
            {
                MessageBox.Show("请选中形状");
            }
        }

        public void selectallshapes()
        {
            PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
            slide.Shapes.SelectAll();
        }

        public void selectallslides()
        {
            PowerPoint.Slides slides = app.ActivePresentation.Slides;
            List<int> slist = new List<int>();
            foreach (PowerPoint.Slide slide in slides)
            {
                slist.Add(slide.SlideNumber);
            }
            slides.Range(slist.ToArray()).Select();
        }

        public void CopySlide()
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides)
            {
                PowerPoint.SlideRange srange = sel.SlideRange;
                srange.Duplicate();
            }
            else
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;
                slide.Duplicate();   
            }
        }

        public void Align_HTiebian()
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                MessageBox.Show("请至少选择两个形状");
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
                if (count == 1)
                {
                    if (range[1].Type == Office.MsoShapeType.msoGroup)
                    {
                        for (int i = 1; i <= range[1].GroupItems.Count; i++)
                        {
                            for (int j = 2; j <= range[1].GroupItems.Count; j++)
                            {
                                range[1].GroupItems[j].Top = range[1].GroupItems[j - 1].Top;
                                range[1].GroupItems[j].Left = range[1].GroupItems[j - 1].Left + range[1].GroupItems[j - 1].Width;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("请至少选择两个形状");
                    }
                }
                else
                {
                    for (int i = 2; i <= count; i++)
                    {
                        range[i].Left = range[i - 1].Left + range[i - 1].Width;
                        range[i].Top = range[i - 1].Top;
                    }
                }
            }
        }

        public void Align_VTiebian()
        {
            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                MessageBox.Show("请至少选择两个形状");
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
                if (count == 1)
                {
                    if (range[1].Type == Office.MsoShapeType.msoGroup)
                    {
                        for (int i = 1; i <= range[1].GroupItems.Count; i++)
                        {
                            for (int j = 2; j <= range[1].GroupItems.Count; j++)
                            {
                                range[1].GroupItems[j].Top = range[1].GroupItems[j - 1].Top + range[1].GroupItems[j - 1].Height;
                                range[1].GroupItems[j].Left = range[1].GroupItems[j - 1].Left;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("请至少选择两个形状");
                    }
                }
                else
                {
                    for (int i = 2; i <= count; i++)
                    {
                        range[i].Top = range[i - 1].Top + range[i - 1].Height;
                        range[i].Left = range[i - 1].Left;
                    }
                }
            }
        }

        public void Notepad()
        {
            System.Diagnostics.ProcessStartInfo Info = new System.Diagnostics.ProcessStartInfo();
            Info.FileName = "notepad.exe ";
            System.Diagnostics.Process Proc = System.Diagnostics.Process.Start(Info);
        }

        private void 静音ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.CommandSound == true)
            {
                Properties.Settings.Default.CommandSound = false;
                静音ToolStripMenuItem.Text = "取消静音";
                Properties.Settings.Default.Save();
            }
            else
            {
                Properties.Settings.Default.CommandSound = true;
                静音ToolStripMenuItem.Text = "静音";
                Properties.Settings.Default.Save();
            }
        }

    }
}
