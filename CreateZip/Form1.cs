using Microsoft.Win32;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Aspose.Cells;
using System.Collections.Generic;
using System.Drawing;

namespace CreateZip
{
    public partial class Form1 : Form
    {
        private static string connStr = "User Id=NCHOME;Password=NCHOME;Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.88.165)(PORT=1521)))(CONNECT_DATA=(SERVICE_NAME=ORCL)))"; //orcale数据库连接字符串

//        private string sqlStr = @"select 
//org_group.name as groupName,
//org_orgs.name as orgName, 
//wa_waclass.name as waName, --工资表名字 
//wa_data.cyear as cyear, --年份
//wa_data.cperiod as cperiod, --月份
//wa_data.checkflag as checkflag, --是否审核
//wa_data.cyearperiod as cyearperiod, --工资年度和月份
//bd_psndoc.name as psnname,
//bd_psndoc.id as id,
//wa_data.pk_psndoc as psnpk,
//wa_data.c_1 as sfzh,--身份证号
//wa_data.c_2 as zwjb,--专业技术职务级别
//wa_data.c_3 as khqk,--考核情况
//wa_data.c_4 as zyzgjsdj,--职业资格技术等级
//wa_data.c_5 as gwxz,--岗位性质
//wa_data.c_6 as xb,--性别
//wa_data.c_7 as dw,--单位
//wa_data.c_8 as xm,--项目
//nvl(wa_data.f_1,0) as yfhj, --应发合计
//nvl(wa_data.f_2,0) as kkhj, --扣款合计
//nvl(wa_data.f_3,0) as sfhj, --实发合计
//nvl(wa_data.f_4,0) as bcksjs, --本次扣税基数
//nvl(wa_data.f_5,0) as bcks, --本次扣税
//nvl(wa_data.f_6,0) as fksjs, --已扣税基数
//nvl(wa_data.f_7,0) as yks, --已扣税
//nvl(wa_data.f_8,0) as bfje, --补发金额
//nvl(wa_data.f_9,0) as bfks, --补发扣税
//nvl(wa_data.f_15,0) + nvl(wa_data.f_43,0)as gwgz, --岗位工资
//nvl(wa_data.f_16,0) as tjgz, --调节工资
//nvl(wa_data.f_17,0) as ydjxgz, --月度业绩工资
//nvl(wa_data.f_18,0) as zyjsjt, --专业技术津贴
//nvl(wa_data.f_19,0) as bcks, --专业技能津贴
//nvl(wa_data.f_20,0) as zhxbt, --综合性补贴
//nvl(wa_data.f_21,0) as bmbt, --保密补贴
//nvl(wa_data.f_22,0) as yjrcjt, --引进人才津贴
//nvl(wa_data.f_23,0) as bfks, --少数民族（回族）补贴
//nvl(wa_data.f_24,0) as jjjcbarybt, --纪检监察办案人员补贴
//nvl(wa_data.f_25,0) as xcbz, --行车补助
//nvl(wa_data.f_26,0) as xcjt, --现场津贴
//nvl(wa_data.f_28,0) as qtjbt, --其他津补贴
//nvl(wa_data.f_29,0)+nvl(wa_data.f_56,0) as jbgz, --加班工资
//nvl(wa_data.f_30,0) as nbtygwgz, --内部退养岗位工资
//nvl(wa_data.f_31,0) as nbtytjgz, --内部退养调节工资
//nvl(wa_data.f_32,0) as syf, --事业费
//nvl(wa_data.f_33,0) as wsjt, --卫生津贴
//nvl(wa_data.f_34,0) as jhnjt, --教护龄津贴
//nvl(wa_data.f_36,0) as ntzhxbt, --内退综合性补贴
//nvl(wa_data.f_37,0) as jbnx, --基本年薪
//nvl(wa_data.f_38,0) as dgryshf, --待岗人员生活费
//nvl(wa_data.f_39,0) as xgryshf, --下岗人员生活费
//nvl(wa_data.f_40,0) as cbazshf, --病残安置生活费
//nvl(wa_data.f_42,0) as jbxj, --基本薪金
//nvl(wa_data.f_44,0) as xjgz, --薪级工资
//nvl(wa_data.f_45,0) as ntshf, --内退生活费
//nvl(wa_data.f_46,0)+nvl(wa_data.f_89,0) as nxwsr, --年薪外收入
//nvl(wa_data.f_47,0) as gngz, --工龄工资
//nvl(wa_data.f_48,0) as jthz, --津贴汇总
//nvl(wa_data.f_49,0) as bthz, --补贴汇总
//nvl(wa_data.f_50,0) as ydj, --月度奖
//nvl(wa_data.f_51,0) as jdj, --季度奖
//nvl(wa_data.f_52,0) as nzj, --年终奖
//nvl(wa_data.f_53,0) as qtdxj, --其他单项奖
//nvl(wa_data.f_54,0)+nvl(wa_data.f_87,0) as dqjxnx, --当期绩效年薪
//nvl(wa_data.f_55,0)+nvl(wa_data.f_88,0) as yqjxnx, --延期绩效年薪
//nvl(wa_data.f_57,0) as qtxm, --其他项目
//nvl(wa_data.f_58,0) as yfze, --应发总额
//nvl(wa_data.f_69,0)+nvl(wa_data.f_92,0) as ynbxgrjn, --养老保险个人缴纳
//nvl(wa_data.f_70,0)+nvl(wa_data.f_93,0) as sybxgrjn, --失业保险个人缴纳
//nvl(wa_data.f_71,0)+nvl(wa_data.f_94,0) as ylbxgrjn, --医疗保险个人缴纳
//nvl(wa_data.f_72,0)+nvl(wa_data.f_95,0) as zfgjjgrjn, --住房公积金个人缴纳
//nvl(wa_data.f_73,0)+nvl(wa_data.f_96,0) as qynjgrjn, --企业年金个人缴纳
//nvl(wa_data.f_74,0) as bcgjjgrjn, --补充公积金个人缴纳
//nvl(wa_data.f_75,0) as bcylbxgrjn, --补充医疗保险个人缴纳
//nvl(wa_data.f_76,0)+nvl(wa_data.f_86,0) as qtkk, --其他扣款

//nvl(wa_data.f_81,0) as ndyjgz, --年度业绩工资
//nvl(wa_data.f_82,0) as xfjt, --信访津贴
//nvl(wa_data.f_83,0) as jt, --津贴
//nvl(wa_data.f_84,0) as bt, --补贴
//nvl(wa_data.f_90,0) as zdgzjj, --重点工作奖励
//nvl(wa_data.f_91,0) as bfjx, --补发基薪
//nvl(wa_data.f_97,0) as hmbt, --回民补贴
//nvl(wa_data.f_98,0) as bcgz, --补差工资
//nvl(wa_data.f_99,0) as tgrygwgz, --特安人员岗位工资
//nvl(wa_data.f_100,0) as tatjgz, --特安人员调节工资
//nvl(wa_data.f_101,0) as tasyf, --特安人员事业费
//nvl(wa_data.f_102,0) as tajnjt, --特安人员技术技能津贴
//nvl(wa_data.f_103,0) as tajhljt, --特安人员教护龄津贴
//nvl(wa_data.f_104,0) as tawsjt, --特安人员卫生津贴
//nvl(wa_data.f_105,0) as tassmzjt, --特安人员少数民族（回族）津贴
//nvl(wa_data.f_106,0) as tazhbt, --特安人员综合性补贴
//nvl(wa_data.f_108,0) as dxsgz, --大学生工资
//nvl(wa_data.f_109,0) as bz, --备注
//nvl(wa_data.f_110,0) as gsjn, --个税缴纳
//nvl(wa_data.f_111,0) as sfje, --实发金额
//nvl(wa_data.f_112,0) as bfdq, --补发当期
//nvl(wa_data.f_113,0) as bfyq, --补发延期
//nvl(wa_data.f_114,0) as kfjx, --扣发基薪
//nvl(wa_data.f_115,0) as kfdq, --扣发当期
//nvl(wa_data.f_116,0) as kfyq, --扣发延期
//nvl(wa_data.f_117,0) as aqyfxdyjl, --全员安全风险抵押奖励
//nvl(wa_data.f_118,0) as dxjl, --单项奖励
//nvl(wa_data.f_119,0) as zdgzjl, --重点工作奖励（含专项抵押奖励）
//nvl(wa_data.f_120,0) as bftjgz, --补发调节工资
//nvl(wa_data.f_121,0) as bfdxsgz, --补发大学生工资
//nvl(wa_data.f_122,0) as bfzyjsjt, --补发专业技术津贴
//nvl(wa_data.f_123,0) as bfzyjnjt, --补发专业技能津贴
//nvl(wa_data.f_124,0) as bfyjrcjt, --补发引进人才津贴
//nvl(wa_data.f_125,0) as bfzhxbt, --补发综合性补贴
//nvl(wa_data.f_126,0) as bfbmbt, --补发保密补贴
//nvl(wa_data.f_127,0) as bfssmzbt, --补发少数民族（回族）补贴
//nvl(wa_data.f_128,0) as bfjjjcbarybt, --补发纪检监察办案人员补贴
//nvl(wa_data.f_129,0) as bfxcbt, --补发行车补贴
//nvl(wa_data.f_130,0) as bfxfjt, --补发信访津贴
//nvl(wa_data.f_131,0) as bfxcjt, --补发现场津贴
//nvl(wa_data.f_132,0) as kfgwgz, --扣发岗位工资
//nvl(wa_data.f_133,0) as kftjgz, --扣发调节工资
//nvl(wa_data.f_134,0) as kfdxsgz, --扣发大学生工资
//nvl(wa_data.f_135,0) as kfzyjsjt, --扣发专业技术津贴
//nvl(wa_data.f_136,0) as kfzyjnjt, --扣发专业技能津贴
//nvl(wa_data.f_137,0) as kfyjrcjt, --扣发引进人才津贴
//nvl(wa_data.f_138,0) as kfzhbt, --扣发综合性补贴
//nvl(wa_data.f_139,0) as kfbmbt, --扣发保密补贴
//nvl(wa_data.f_140,0) as kfssmzbt, --扣发少数民族（回族）补贴
//nvl(wa_data.f_141,0) as kfjjjcbarybt, --扣发纪检监察办案人员补贴
//nvl(wa_data.f_142,0) as kfxcbt, --扣发行车补贴
//nvl(wa_data.f_143,0) as kfxfjt, --扣发信访津贴
//nvl(wa_data.f_144,0) as kfxcjt, --扣发现场津贴
//case when wa_data.pk_wa_class='1001A2100000000049RO' then 0  else nvl(wa_data.f_145,0) end bf, --补发
//case when wa_data.pk_wa_class='1001A2100000000049RO' then  nvl(wa_data.f_145,0) else 0 end  kfzyjnjt, --特安补发
//case when wa_data.pk_wa_class='1001A2100000000049RO' then 0  else nvl(wa_data.f_146,0) end bf, --扣发
//case when wa_data.pk_wa_class='1001A2100000000049RO' then  nvl(wa_data.f_146,0) else 0 end kfzyjnjt, --特安扣发
//nvl(wa_data.f_147,0) as fsxsphsbzf, --放射性食品伙食补助费
//nvl(wa_data.f_148,0) as tpgz, --聘用工资
//nvl(wa_data.f_152,0) as yzjxnx, --预支绩效年薪
//nvl(wa_data.f_153,0) as tjgztzx, --调节工资调整项
//nvl(wa_data.f_155,0) as dwjnbi, --单位缴纳比例
//nvl(wa_data.f_156,0) as njjs, --年金基数
//nvl(wa_data.f_157,0) as grjnbl, --个人缴纳比例
//nvl(wa_data.f_158,0) as dwjnbl, --单位缴交额
//nvl(wa_data.f_159,0) as grjje, --个人缴交额
//nvl(wa_data.f_160,0) as ylbxqyjn, --养老保险企业缴纳
//nvl(wa_data.f_161,0) as ylbxqyjn, --医疗保险企业缴纳
//nvl(wa_data.f_162,0) as sybxqyjn, --失业保险企业缴纳
//nvl(wa_data.f_163,0) as gsbxqyjn, --工伤保险企业缴纳
//nvl(wa_data.f_164,0) as sybxqyjn, --生育保险企业缴纳
//nvl(wa_data.f_165,0) as zfgjjqyjn, --住房公积金企业缴纳
//nvl(wa_data.f_166,0) as qynjjn, --年金企业缴纳
//nvl(wa_data.f_167,0) as sbqybj, --社保企业补缴
//nvl(wa_data.f_168,0) as sbgrbk --社保个人补扣
//from wa_data
//left outer join org_group on org_group.pk_group =wa_data.pk_group 
//left outer join org_orgs on org_orgs.pk_org=wa_data.pk_org
//left outer join wa_waclass on wa_waclass.pk_wa_class=wa_data.pk_wa_class
//left outer join bd_psndoc on bd_psndoc.pk_psndoc=wa_data.pk_psndoc ";

        private string sqlStr = @"select 
org_group.name as groupName,
org_orgs.name as orgName, 
wa_data.cyear as cyear, --年份
wa_data.cperiod as cperiod, --月份
wa_data.cyearperiod as cyearperiod, --工资年度和月份
bd_psndoc.name as psnname,
bd_psndoc.id as id,--身份证号
wa_data.pk_psndoc as psnpk,
nvl(wa_data.f_15,0) + nvl(wa_data.f_43,0)as gwgz, --岗位工资
nvl(wa_data.f_79,0) as bfgwgz, --补发岗位工资
nvl(wa_data.f_16,0) as tjgz, --调节工资
nvl(wa_data.f_120,0) as bftjgz, --补发调节工资
nvl(wa_data.f_19,0) as bcks, --专业技能津贴
nvl(wa_data.f_123,0) as bfzyjnjt, --补发专业技能津贴
nvl(wa_data.f_18,0) as zyjsjt, --专业技术津贴
nvl(wa_data.f_122,0) as bfzyjsjt, --补发专业技术津贴
nvl(wa_data.f_34,0) as jhnjt, --教护龄津贴
nvl(wa_data.f_23,0) as bfks, --少数民族（回族）补贴
nvl(wa_data.f_20,0) as zhxbt, --综合性补贴
nvl(wa_data.f_125,0) as bfzhxbt, --补发综合性补贴
nvl(wa_data.f_21,0) as bmbt, --保密补贴
nvl(wa_data.f_126,0) as bfbmbt, --补发保密补贴
nvl(wa_data.f_82,0) as xfjt, --信访津贴
nvl(wa_data.f_24,0) as jjjcbarybt, --纪检监察办案人员补贴
nvl(wa_data.f_29,0)+nvl(wa_data.f_56,0) as jbgz, --加班工资
nvl(wa_data.f_17,0) as ydjxgz, --月度业绩工资
nvl(wa_data.f_69,0)+nvl(wa_data.f_92,0) as ynbxgrjn, --养老保险个人缴纳
nvl(wa_data.f_70,0)+nvl(wa_data.f_93,0) as sybxgrjn, --失业保险个人缴纳
nvl(wa_data.f_71,0)+nvl(wa_data.f_94,0) as ylbxgrjn, --医疗保险个人缴纳
nvl(wa_data.f_72,0)+nvl(wa_data.f_95,0) as zfgjjgrjn, --住房公积金个人缴纳
nvl(wa_data.f_73,0)+nvl(wa_data.f_96,0) as qynjgrjn, --企业年金个人缴纳
nvl(wa_data.f_110,0) as gsjn, --个税缴纳
nvl(wa_data.f_111,0) as sfje, --实发金额
nvl(wa_data.f_74,0) as bcgjjgrjn, --补充公积金个人缴纳
nvl(wa_data.f_75,0) as bcylbxgrjn, --补充医疗保险个人缴纳
nvl(wa_data.f_3,0) as sfhj --实发合计
from wa_data
left outer join org_group on org_group.pk_group =wa_data.pk_group 
left outer join org_orgs on org_orgs.pk_org=wa_data.pk_org
left outer join wa_waclass on wa_waclass.pk_wa_class=wa_data.pk_wa_class
left outer join bd_psndoc on bd_psndoc.pk_psndoc=wa_data.pk_psndoc ";

        public Form1()
        {
            this.StartPosition = FormStartPosition.CenterScreen;
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //判断是否安装WinRar
            var registryKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WinRAR.exe");
            if (registryKey == null)
            {
                MessageBox.Show("请确认服务器上已安装WinRar应用！");
            }
            BindCombox();
        }

        #region 提交按钮触发事件
        /// <summary>
        /// 点击提交事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            LoadingHelper.ShowLoading("努力生成中，请稍后...", this, o =>
            {
                string root = Application.StartupPath + "\\Salary\\";
                if (Directory.Exists(root) == false)
                {
                    Directory.CreateDirectory(root);
                }
                string sqlString = "Select name,code,id,mobile,sex From bd_psndoc";
                DataTable dt = ExecuteDataTable(sqlString, new OracleParameter());
                ExceloutByAspose(dt, System.AppDomain.CurrentDomain.BaseDirectory + "/Salary/bd_psndoc.xls");

                string salarySql = sqlStr;
                List<OracleParameter> ilistStr = new List<OracleParameter>();
                List<string> whereList = new List<string>();
                if (dateTimePicker1.Value != null)
                {
                    ilistStr.Add(new OracleParameter(":year", dateTimePicker1.Value.ToString("yyyyMM")));
                    whereList.Add(" wa_data.cyearperiod = :year ");
                }
                if (!string.IsNullOrEmpty(comboBox1.SelectedValue.ToString()))
                {
                    ilistStr.Add(new OracleParameter(":org", comboBox1.SelectedValue.ToString()));
                    whereList.Add(" wa_data.pk_org = :org ");
                }
                if (!string.IsNullOrEmpty(comboBox2.SelectedValue.ToString()))
                {
                    ilistStr.Add(new OracleParameter(":waclass", comboBox2.SelectedValue.ToString()));
                    whereList.Add(" wa_data.pk_wa_class = :waclass ");
                }
                if (whereList.Count > 0)
                {
                    salarySql += " where " + string.Join(" and ", whereList);
                }
                OracleParameter[] parameters = ilistStr.ToArray();
                DataTable dt2 = ExecuteDataTable(salarySql, parameters);
                ExceloutByAspose(dt2, System.AppDomain.CurrentDomain.BaseDirectory + "/Salary/salary.xls");
                CreateRarFile();
            });
        }
        #endregion

        #region 选择目录按钮触发事件
        /// <summary>
        /// 选择生成的压缩文件存放路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = SelectPath();
        }
        #endregion

        #region 生成Excel文件
        /// <summary>
        /// 生成Excel文件(StreamWriter)
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="path"></param>
        private void Excelout(DataTable ds, string path)
        {
            try
            {
                long totalCount = ds.Rows.Count;
                long rowRead = 0;
                float percent = 0;
                StreamWriter sw = new StreamWriter(path, false, Encoding.GetEncoding("gb2312"));
                StringBuilder sb = new StringBuilder();
                for (int k = 0; k < ds.Columns.Count; k++)
                {
                    sb.Append(ds.Columns[k].ColumnName.ToString() + "\t");
                }
                sb.Append(Environment.NewLine);
                for (int i = 0; i < ds.Rows.Count; i++)
                {
                    rowRead++;
                    percent = ((float)(100 * rowRead)) / totalCount;
                    Application.DoEvents();
                    for (int j = 1; j < ds.Columns.Count; j++)
                    {
                        sb.Append(ds.Rows[i][j].ToString() + "\t");
                    }
                    sb.Append(Environment.NewLine);
                }
                sw.Write(sb.ToString());
                sw.Flush();
                sw.Close();
                CreateRarFile();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 生成Excel文件(Aspose.Cells)
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="path"></param>
        private void ExceloutByAspose(DataTable ds, string path)
        {
            try
            {
                License li = new License();
                Workbook wk = new Workbook();
                Worksheet ws = wk.Worksheets[0];
                for (int k = 0; k < ds.Columns.Count; k++)
                {
                    ws.Cells[0, k].PutValue(ds.Columns[k].ColumnName.ToString());
                }
                for (int i = 0; i < ds.Rows.Count; i++)
                {
                    for (int j = 0; j < ds.Columns.Count; j++)
                    {
                        ws.Cells[i + 1, j].PutValue(ds.Rows[i][j].ToString());
                    }
                }
                wk.Save(path);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region 创建压缩文件
        /// <summary>
        /// 创建压缩文件
        /// </summary>
        private void CreateRarFile()
        {
            if (Directory.Exists(textBox1.Text) == false)//如果不存在就创建file文件夹
            {
                Directory.CreateDirectory(textBox1.Text);
            }
            string strzipPath = textBox1.Text + @"\" + DateTime.Now.ToString("yyyy-MM-dd") + "工资表 ";
            string strtxtPath = "Salary";
            string password = "123456";
            bool flag = Zip(strzipPath, strtxtPath, password);
            if (flag)
            {
                if (Directory.Exists(Application.StartupPath + "\\Salary\\"))
                {
                    Directory.Delete(Application.StartupPath + "\\Salary\\", true);
                }
                MessageBox.Show("文件生成成功!");
            }
            else
            {
                MessageBox.Show("文件生成失败!");
            }
        }
        #endregion

        #region 弹出选择目录对话框
        /// <summary>
        /// 弹出一个选择目录的对话框
        /// </summary>
        private string SelectPath()
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.ShowDialog();
            return path.SelectedPath;
        }
        #endregion

        #region 生成加密压缩文件方法
        /// <summary>
        /// 加密压缩
        /// </summary>
        /// <param name="strzipPath">压缩包路径</param>
        /// <param name="strtxtPath">待压缩的文件路径</param>
        /// <param name="password">加密密码</param>
        /// <returns></returns>
        public bool Zip(string strzipPath, string strtxtPath, string password)
        {
            try
            {
                Process Process1 = new Process();
                Process1.StartInfo.FileName = "Winrar.exe";
                Process1.StartInfo.CreateNoWindow = true;
                Process1.StartInfo.Arguments = " a -afzip -p" + password + " " + strzipPath + strtxtPath;
                Process1.Start();
                Process1.WaitForExit();
                if (Process1.HasExited)
                {
                    Process1.WaitForExit();
                    return true;
                }
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }
        #endregion

        #region 执行SQL语句,返回DataTable;
        /// <summary>
        /// 数据库数据查询返回DataTable
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static DataTable ExecuteDataTable(string sql, params OracleParameter[] parameters)
        {
            using (OracleConnection conn = new OracleConnection(connStr))
            {
                conn.Open();
                using (OracleCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = sql;
                    cmd.Parameters.AddRange(parameters);
                    OracleDataAdapter adapter = new OracleDataAdapter(cmd);
                    DataTable datatable = new DataTable();
                    adapter.Fill(datatable);
                    return datatable;
                }
            }
        }
        #endregion

        #region 组织部门、方案下拉框数据绑定
        private void BindCombox()
        {
            string orgSqlString = "Select pk_org as id,name From org_orgs order by name";
            DataTable orgdt = ExecuteDataTable(orgSqlString, new OracleParameter());
            DataRow dr1 = orgdt.NewRow();
            dr1["id"] = "";
            dr1["name"] = "请选择";
            orgdt.Rows.InsertAt(dr1, 0);
            comboBox1.DataSource = orgdt;
            comboBox1.ValueMember = "id";
            comboBox1.DisplayMember = "name";

            string waclassSqlString = "Select pk_wa_class as id,name From wa_waclass order by name";
            DataTable waclassdt = ExecuteDataTable(waclassSqlString, new OracleParameter());
            DataRow dr = waclassdt.NewRow();
            dr["id"] = "";
            dr["name"] = "请选择";
            waclassdt.Rows.InsertAt(dr, 0);
            comboBox2.DataSource = waclassdt;
            comboBox2.ValueMember = "id";
            comboBox2.DisplayMember = "name";
           
        }

        #endregion

    }
}