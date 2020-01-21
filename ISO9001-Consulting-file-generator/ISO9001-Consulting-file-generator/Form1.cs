using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void FindAndReplace(Word.Application wordApp, object ToFindText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref ToFindText, ref matchCase, ref matchWholeWord, ref matchWildCards,
                                            ref matchSoundLike, ref nmatchAllforms, ref forward, ref wrap, ref format,
                                            ref replaceWithText, ref replace, ref matchKashida, ref matchDiactitics,
                                            ref matchAlefHamza, ref matchControl);
        }

        private void CreateWordDocument(object filename, object SaveAs)
        {
            Word.Application wordApp = new Word.Application();
            object missing = System.Reflection.Missing.Value;
            Word.Document myWordDoc = null;

            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref filename);
                myWordDoc.Activate();

                this.FindAndReplace(wordApp, "体系审核年份A", textbox1.Text);
                this.FindAndReplace(wordApp, "企业名称A", textBox2.Text);
                this.FindAndReplace(wordApp, "企业注册地址A", textBox3.Text);
                this.FindAndReplace(wordApp, "企业审核地址A", textBox4.Text);
                this.FindAndReplace(wordApp, "联系人A", textBox5.Text);
                this.FindAndReplace(wordApp, "企业联系电话A", textBox6.Text);
                this.FindAndReplace(wordApp, "公司简介A", textBox7.Text);
                this.FindAndReplace(wordApp, "企业外部宗旨A", textBox8.Text);
                this.FindAndReplace(wordApp, "企业内部宗旨A", textBox9.Text);
                this.FindAndReplace(wordApp, "企业战略方向A", textBox10.Text);
                this.FindAndReplace(wordApp, "企业产品A", textBox11.Text);
                this.FindAndReplace(wordApp, "质量手册文件编号A", textBox12.Text);
                this.FindAndReplace(wordApp, "质量手册版本A", textBox13.Text);
                this.FindAndReplace(wordApp, "质量手册编制人A", textBox14.Text);
                this.FindAndReplace(wordApp, "质量手册审核人A", textBox15.Text);
                this.FindAndReplace(wordApp, "质量手册批准人A", textBox16.Text);
                this.FindAndReplace(wordApp, "质量手册发布时间A", textBox17.Text);
                this.FindAndReplace(wordApp, "质量手册实施时间A", textBox18.Text);
                this.FindAndReplace(wordApp, "体系运行开始时间A", textBox19.Text);
                this.FindAndReplace(wordApp, "质量手册修订时间A", textBox20.Text);
                this.FindAndReplace(wordApp, "质量手册覆盖范围A", textBox21.Text);
                this.FindAndReplace(wordApp, "外包过程A", textBox22.Text);
                this.FindAndReplace(wordApp, "质量方针A", textBox23.Text);
                this.FindAndReplace(wordApp, "质量目标A", textBox24.Text);
                this.FindAndReplace(wordApp, "质量目标B", textBox25.Text);
                this.FindAndReplace(wordApp, "总经理A", textBox26.Text);
                this.FindAndReplace(wordApp, "管理者代表A", textBox27.Text);
                this.FindAndReplace(wordApp, "产品实现部门A", textBox28.Text);
                this.FindAndReplace(wordApp, "产品实现部门负责人A", textBox29.Text);
                this.FindAndReplace(wordApp, "产品质量控制部门A", textBox30.Text);
                this.FindAndReplace(wordApp, "产品质量控制部门负责人A", textBox31.Text);
                this.FindAndReplace(wordApp, "采购销售人事管理部门A", textBox32.Text);
                this.FindAndReplace(wordApp, "采购销售人事管理部门负责人A", textBox33.Text);
                this.FindAndReplace(wordApp, "财务部负责人A", textBox34.Text);
                this.FindAndReplace(wordApp, "相关方识别时间A", textBox35.Text);
                this.FindAndReplace(wordApp, "组织内外部环境因素识别时间A", textBox36.Text);
                this.FindAndReplace(wordApp, "风险和机遇评估时间A", textBox37.Text);
                this.FindAndReplace(wordApp, "风险和机遇评估实施时间A", textBox38.Text);
                this.FindAndReplace(wordApp, "风险和机遇评价时间A", textBox39.Text);
                this.FindAndReplace(wordApp, "内审时间1A", textBox40.Text);
                this.FindAndReplace(wordApp, "内审时间2A", textBox41.Text);
                this.FindAndReplace(wordApp, "内审计划编制时间A", textBox42.Text);
                this.FindAndReplace(wordApp, "内审组长A", textBox43.Text);
                this.FindAndReplace(wordApp, "内审组员A", textBox44.Text);
                this.FindAndReplace(wordApp, "内审不符合责任部门A", textBox45.Text);
                this.FindAndReplace(wordApp, "内审不符合责任部门负责人A", textBox46.Text);
                this.FindAndReplace(wordApp, "内审不符合条款A", textBox47.Text);
                this.FindAndReplace(wordApp, "内审不符合纠正措施完成时间A", textBox48.Text);
                this.FindAndReplace(wordApp, "内审不符合整改确认时间A", textBox49.Text);
                this.FindAndReplace(wordApp, "内审不符合整改验证时间A", textBox50.Text);
                this.FindAndReplace(wordApp, "内审不符合描述A", textBox51.Text);
                this.FindAndReplace(wordApp, "管理评审实施时间A", textBox52.Text);
                this.FindAndReplace(wordApp, "管理评审计划时间A", textBox53.Text);
                this.FindAndReplace(wordApp, "管理评审改进项目A", textBox54.Text);
                this.FindAndReplace(wordApp, "管理评审改进计划时间A", textBox55.Text);
                this.FindAndReplace(wordApp, "管理评审改进措施A", textBox56.Text);
                this.FindAndReplace(wordApp, "管理评审改进提出者A", textBox57.Text);
                this.FindAndReplace(wordApp, "管理评审改进责任人A", textBox58.Text);
                this.FindAndReplace(wordApp, "管理评审改进跟进人A", textBox59.Text);
                this.FindAndReplace(wordApp, "管理评审改进完成时间A", textBox60.Text);
                this.FindAndReplace(wordApp, "培训内容1时间A", textBox61.Text);
                this.FindAndReplace(wordApp, "培训内容1实施时间A", textBox62.Text);
                this.FindAndReplace(wordApp, "培训内容2时间A", textBox63.Text);
                this.FindAndReplace(wordApp, "培训内容2实施时间A", textBox64.Text);
                this.FindAndReplace(wordApp, "培训内容3时间A", textBox65.Text);
                this.FindAndReplace(wordApp, "培训内容3实施时间A", textBox66.Text);
                this.FindAndReplace(wordApp, "培训内容4时间A", textBox67.Text);
                this.FindAndReplace(wordApp, "培训内容4实施时间A", textBox68.Text);
                this.FindAndReplace(wordApp, "培训内容5时间A", textBox69.Text);
                this.FindAndReplace(wordApp, "培训内容5实施时间A", textBox70.Text);
                this.FindAndReplace(wordApp, "培训内容6时间A", textBox71.Text);
                this.FindAndReplace(wordApp, "培训内容6实施时间A", textBox72.Text);
                this.FindAndReplace(wordApp, "培训内容7时间A", textBox73.Text);
                this.FindAndReplace(wordApp, "培训内容7实施时间A", textBox74.Text);
            }
            else
            {
                MessageBox.Show("File not found!");
            }
            

            //save as
            myWordDoc.SaveAs2(ref SaveAs);
            myWordDoc.Close();
            wordApp.Quit();
            MessageBox.Show("File Created!");
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_click(Object sender, EventArgs e)
        {
            //C:\Users\rayji\OneDrive\桌面\20200109张家港宏威——整合版.doc
            //C:\Users\rayji\OneDrive\桌面\20200109张家港宏威——完成版.doc
            CreateWordDocument(textBox75.Text, textBox76.Text);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox11_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox12_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox13_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox14_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox15_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox16_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox17_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox18_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox19_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox20_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox21_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox22_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox23_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox24_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox25_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox26_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox27_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox28_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox29_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox30_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox31_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox32_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox33_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox34_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox35_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox36_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox37_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox38_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox39_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox40_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox41_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox42_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox43_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox44_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox45_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox46_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox47_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox48_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox49_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox50_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox51_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox52_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox53_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox54_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox55_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox56_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox57_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox58_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox59_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox60_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox61_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox62_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox63_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox64_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox65_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox66_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox67_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox68_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox69_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox70_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox71_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox72_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox73_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox74_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox75_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox76_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
