using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using SchoolHealth.Properties;
using DevComponents.DotNetBar.Controls;

namespace SchoolHealth
{
    public partial class RecordInputControl : UserControl
    {
        List<ExpandablePanel> expandPanels;
        public bool IsNewRecordAddedButNotViewHistory {get;set;}
        bool isCanProcessNewAndSave = true;

        public SchoolHealthRecord Record { get; set; }
        bool isUpdateMode;
        public bool IsUpdateMode
        {
            get { return isUpdateMode; }
            set
            {
                isUpdateMode = value;
                if (isUpdateMode)
                {
                    panelEx1.Visible = false;
                    foreach (var ep in this.expandPanels)
                    {
                        ep.Enabled = true;
                        btnSaveAndNewEx.Text = "保存修改(F4)";
                        btnSaveEx.Visible = false;
                        btnNewEx.Visible = false;
                        btnSaveAndNewEx.Enabled = true;
                    }
                }
            }
        }

        public RecordInputControl()
        {
            InitializeComponent();
        }

        public void Initialize()
        {
            InitializeComboBoxSource();
            InitInputValue();

            Record = new SchoolHealthRecord();
            expandPanels = new List<ExpandablePanel>()
            {
                expandablePanel1,expandablePanel2,expandablePanel3,expandablePanel4,
                expandablePanel6,expandablePanel7,expandablePanel9,expandablePanel5
            };
            switchButtonExpandCollapse.Value = true;
            switchButtonExpandCollapse_ValueChanged(switchButtonExpandCollapse, null);
            comboBoxExXueXiao.Focus();

            foreach (var ep in expandPanels)
            {
                ep.Enabled = false;
            }

            UpdateBloodPressureUnit();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            groupPanelRecord.ScrollControlIntoView(btnNew);

            foreach (var ep in expandPanels)
            {
                ep.Enabled = true;
            }

            btnNew.Enabled = false;
            btnNewEx.Enabled = false;
            btnSave.Enabled = true;
            btnSaveEx.Enabled = true;
            btnSaveAndNew.Enabled = true;
            btnSaveAndNewEx.Enabled = true;


            if (!switchButtonExpandCollapse.Value)
            {
                switchButtonExpandCollapse.Value = true;
            }
            else
            {
                foreach (var ep in expandPanels)
                {
                    ep.Expanded = true;
                }
            }


            //set default value
            ResetComboBoxSelection();
            InitInputValue();




            if (!Settings.Default.DefaultSchool.IsNullOrEmpty())
            {
                comboBoxExXueXiao.Text = Settings.Default.DefaultSchool;
            
            }
            if (!Settings.Default.DefaultGrade.IsNullOrEmpty())
            {
                //comboBoxExNianJi.Text = Settings.Default.DefaultGrade;
                comboBoxExNianJi.SelectItem(Settings.Default.DefaultGrade);
            }

            if (!Settings.Default.DefaultClass.IsNullOrEmpty())
            {
                comboBoxExBanJi.Text = Settings.Default.DefaultClass;
            }
            if (Settings.Default.CheckDate.Year >= 1990)
            {
                integerInputCheckDateYear.Value = Settings.Default.CheckDate.Year;
                integerInputCheckDateMonth.Value = Settings.Default.CheckDate.Month;
                integerInputCheckDateDay.Value = Settings.Default.CheckDate.Day;
            }
            if (!Settings.Default.Checker.IsNullOrEmpty())
            {
                comboBoxExChecker.Text = Settings.Default.Checker;
            }


            comboBoxExXueXiao.Focus();
        }

        private void btnNewEx_Click(object sender, EventArgs e)
        {
           
            btnNew_Click(null, null);
        }



        private void btnSaveAndNew_Click(object sender, EventArgs e)
        {

            if(IsUpdateMode)
            {
                if (Record.ID == 0) return;

                if(UpdateUIInputToDataModel())
                {
                    if(DBHelper.Update(Record))
                    {
                        Utility.ShowInformation("记录更新成功");
                    }
                }

            }
            else 
            {
                isCanProcessNewAndSave = true;
                btnSave_Click(null, null);
                if (!isCanProcessNewAndSave) return;
                btnNewEx_Click(null, null);
            }
        }

        private void switchButtonExpandCollapse_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                groupPanelRecord.SuspendLayout();
                foreach (var ep in expandPanels)
                {
                    ep.Expanded = switchButtonExpandCollapse.Value;
                }
            }
            catch
            {

            }
            finally
            {
                groupPanelRecord.ResumeLayout(false);
                groupPanelRecord.PerformLayout();
            }

        }
    

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (UpdateUIInputToDataModel())
            {
                btnNew.Enabled = true;
                btnNewEx.Enabled = true;
                btnSave.Enabled = false;
                btnSaveEx.Enabled = false;
                btnSaveAndNew.Enabled = false;
                btnSaveAndNewEx.Enabled = false;
                DBHelper.Add(Record);
         //       isNewRecordAddedButNotViewHistory = true;
            }
            else
            {
                isCanProcessNewAndSave = false;
            }
        }


        private void ResetComboBoxSelection()
        {

            foreach (var gp in this.expandPanels)
            {
                foreach (Control child in gp.Controls)
                {
                    ComboBoxEx cbx = child as ComboBoxEx;
                    if (cbx != null && cbx.DropDownStyle == ComboBoxStyle.DropDownList)
                    {
                        cbx.SelectedIndex = cbx.Items.Count > 0 ? 0 : -1;
                    }
                }
            }
        }

        private void InitializeComboBoxSource()
        {
            ConfigurationHelper.BindDataSource(ConfigurationHelper.年级, comboBoxExNianJi);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.性别, comboBoxExXingBie);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.余杭户籍, comboBoxExYuHang);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.心, comboBoxExXin);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.肺, comboBoxExFei);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.肝, comboBoxExGan);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.脾, comboBoxExPi);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.沙眼, comboBoxExShaYan);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.结膜炎, comboBoxExJieMoYan);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.色觉, comboBoxExSeJue);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.牙周, comboBoxExYaZhou);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.耳, comboBoxExEr);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.鼻, comboBoxExBi);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.扁桃体, comboBoxExBianTaoTi);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.头部, comboBoxExTou);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.颈部, comboBoxExJing);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.胸部, comboBoxExXiong);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.脊柱, comboBoxExJiZhu);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.四肢, comboBoxExSiZhi);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.皮肤, comboBoxExPiFu);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.淋巴结, comboBoxExLinBaJie);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.血型, comboBoxExXueXing);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.蛔虫卵, comboBoxExHuiChongLuan);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.卡疤, comboBoxExKaBa);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.结核菌素试验, comboBoxExJieHeJun);

            ConfigurationHelper.BindDataSource(ConfigurationHelper.谷丙转氨酶, comboBoxExZhuanAnMei);
            ConfigurationHelper.BindDataSource(ConfigurationHelper.胆红素, comboBoxExDanHongSu);

            //ConfigurationHelper.BindDataSource(ConfigurationHelper.年级, textBoxDropDownGrade, RefreshClass);
            // ConfigurationHelper.BindDataSource(ConfigurationHelper.Excel数据类型, comboBoxExExcelExportType);


            //打印机
            //comboBoxExPrint.SelectedIndexChanged -= comboBoxExPrint_SelectedIndexChanged;
            //comboBoxExPrint.DataSource = Utility.LoadPrinters();
            //var dp = Properties.Settings.Default.DefaultPrinter;
            //if (!dp.IsNullOrEmpty())
            //{
            //    var index = comboBoxExPrint.FindStringExact(dp);
            //    if (index > -1)
            //    {
            //        comboBoxExPrint.SelectedIndex = index;
            //        comboBoxExPrint_SelectedIndexChanged(comboBoxExPrint, null);
            //    }
            //}
            //else if (comboBoxExPrint.Items.Count > 0)
            //{
            //    comboBoxExPrint_SelectedIndexChanged(comboBoxExPrint, null);
            //}
            //comboBoxExPrint.SelectedIndexChanged += comboBoxExPrint_SelectedIndexChanged;

            //UI Style
            //comboBoxExUIStyle.SelectedIndexChanged -= comboBoxExUIStyle_SelectedIndexChanged;
            //comboBoxExUIStyle.DataSource = Enum.GetNames(typeof(eStyle));
            //var ds = Properties.Settings.Default.DefaultStyle;
            //if (!ds.IsNullOrEmpty())
            //{
            //    var index = comboBoxExUIStyle.FindStringExact(ds);
            //    if (index > -1)
            //    {
            //        comboBoxExUIStyle.SelectedIndex = index;
            //        comboBoxExUIStyle_SelectedIndexChanged(comboBoxExUIStyle, null);
            //    }
            //}
            //else if (comboBoxExUIStyle.Items.Count > 0)
            //{
            //    comboBoxExUIStyle_SelectedIndexChanged(comboBoxExUIStyle, null);
            //}
            //comboBoxExUIStyle.SelectedIndexChanged += comboBoxExUIStyle_SelectedIndexChanged;

        }

        private void InitInputValue()
        {
            integerInputAYouShang.Value =
                integerInputAYouXia.Value =
                integerInputAZuoShang.Value =
                integerInputAZuoXia.Value = 0;


            integerInputBYouShang.Value = 0;
            integerInputBYouXia.Value = 0;
            integerInputBZuoShang.Value = 0;
            integerInputBZuoXia.Value = 0;

            integerInputDYouShang.Value = 0;
            integerInputDYouXia.Value = 0;
            integerInputDZuoShang.Value = 0;
            integerInputDZuoXia.Value = 0;

            integerInputFYouShang.Value = 0;
            integerInputFYouXia.Value = 0;
            integerInputFZuoShang.Value = 0;
            integerInputFZuoXia.Value = 0;

            doubleInputHeight.Value = 0;
            doubleInputWeight.Value = 0;
            integerInputLungCapacity.Value = 0;
            doubleInputSuZhangYa.Value = 0;
            doubleInputSouSuoYa.Value = 0;
            integerInputMaiBo.Value = 0;

            doubleInputXueHongDanBai.Value = 0;

            integerInputYiJing.ValueObject = null;
            integerInputChuChao.ValueObject = null;

            doubleInputHeight.ValueObject = null;
            doubleInputWeight.ValueObject = null;
            integerInputLungCapacity.ValueObject = null;
            doubleInputSuZhangYa.ValueObject = null;
            doubleInputSouSuoYa.ValueObject = null;
            integerInputMaiBo.ValueObject = null;
            doubleInputXiongWei.ValueObject = null;
            doubleInputYaoWei.ValueObject = null;
            doubleInputTunWei.ValueObject = null;

            doubleInputXueHongDanBai.ValueObject = null;





            textBoxXXingMing.Text = string.Empty;
            textBoxXXueHao.Text = string.Empty;
            doubleInputZuoYan.Value = 5.0;
            doubleInputYouYan.Value = 5.0;


            integerInputBirthDateMonth.ValueObject = null;
            integerInputBirthDateDay.ValueObject = null;

            integerInputSickDateYear.ValueObject = null;
            integerInputSickDateMonth.ValueObject = null;
            integerInputSickDateDay.ValueObject = null;
        }

        private void ScrollToVisible(Control ctrl)
        {
            groupPanelRecord.ScrollControlIntoView(ctrl);
            ctrl.Focus();
        }

        private bool UpdateUIInputToDataModel()
        {


            //check

            if (comboBoxExXueXiao.Text.IsNullOrEmpty())
            {
                Utility.ShowInformation("学校不可为空!", this);
                ScrollToVisible(comboBoxExXueXiao);
                return false;
            }

            if (comboBoxExBanJi.Text.IsNullOrEmpty())
            {
                Utility.ShowInformation("班级不可为空!", this);
                ScrollToVisible(comboBoxExBanJi);
                return false;
            }

            if (textBoxXXingMing.Text.IsNullOrEmpty())
            {
                Utility.ShowInformation("姓名不可为空!", this);
                ScrollToVisible(textBoxXXingMing);
                return false;
            }

            if (textBoxXXueHao.Text.IsNullOrEmpty())
            {
                Utility.ShowInformation("学号不可为空!", this);
                ScrollToVisible(textBoxXXueHao);
                return false;
            }

            if (comboBoxExChecker.Text.IsNullOrEmpty())
            {
                Utility.ShowInformation("体检单位不可为空!", this);
                ScrollToVisible(comboBoxExChecker);
                return false;
            }

            if ((int)doubleInputHeight.Value == 0 || (int)doubleInputHeight.Value == (int)doubleInputHeight.MinValue)
            {
                Utility.ShowInformation("请输入正确的身高!", this);
                ScrollToVisible(doubleInputHeight);
                return false;
            }

            if ((int)doubleInputWeight.Value == 0 || (int)doubleInputWeight.Value == (int)doubleInputWeight.MinValue)
            {
                Utility.ShowInformation("请输入正确的体重!", this);
                ScrollToVisible(doubleInputWeight);
                return false;
            }

            if ((int)doubleInputXiongWei.Value == 0 || (int)doubleInputXiongWei.Value == (int)doubleInputXiongWei.MinValue)
            {
                Utility.ShowInformation("请输入正确的胸围!", this);
                ScrollToVisible(doubleInputWeight);
                return false;
            }

            //if (integerInputLungCapacity.Value == 0)
            //{
            //    Utility.ShowInformation("请输入肺活量!", this);
            //    ScrollToVisible(integerInputLungCapacity);
            //    return false;
            //}

            if (doubleInputSuZhangYa.Value == 0 || doubleInputSuZhangYa.Value == doubleInputSuZhangYa.MinValue)
            {
                Utility.ShowInformation("请输入正确的舒张压!", this);
                ScrollToVisible(doubleInputSuZhangYa);
                return false;
            }

            if (doubleInputSouSuoYa.Value == 0 || doubleInputSouSuoYa.Value == doubleInputSouSuoYa.MinValue)
            {
                Utility.ShowInformation("请输入正确的收缩压!", this);
                ScrollToVisible(doubleInputSouSuoYa);
                return false;
            }

            if (integerInputMaiBo.Value == 0 || integerInputMaiBo.Value == integerInputMaiBo.MinValue)
            {
                Utility.ShowInformation("请输入正确的脉搏!", this);
                ScrollToVisible(integerInputMaiBo);
                return false;
            }


            //if ((int)doubleInputXueHongDanBai.Value == 0)
            //{
            //    Utility.ShowInformation("请输入血红蛋白值!", this);
            //    ScrollToVisible(doubleInputXueHongDanBai);
            //    return false;
            //}


            //assign 
            //1) 基本信息
            Record.XueXiao = comboBoxExXueXiao.Text;
            Record.BanJi = comboBoxExBanJi.Text;
            Record.NianJi = comboBoxExNianJi.Text;
            Record.XingMing = textBoxXXingMing.Text;
            Record.StringReserved3 = textBoxXXueHao.Text;//学号
            Record.XingBie = comboBoxExXingBie.SelectedIndex == 0;
            Record.ShengRi = new DateTime(integerInputBirthDateYear.Value, integerInputBirthDateMonth.Value, integerInputBirthDateDay.Value);
            Record.YuHangHuJi = comboBoxExYuHang.SelectedIndex == 0;
            Record.TiJianRiQi = new DateTime(integerInputCheckDateYear.Value, integerInputCheckDateMonth.Value, integerInputCheckDateDay.Value);
            Record.TiJianDanWei = comboBoxExChecker.Text;

            //2) 既往病史
            Record.GanYan = checkBoxXGanSick.Checked;
            Record.FeiJieHe = checkBoxXFeiSick.Checked;
            Record.XinZangBing = checkBoxXXinSick.Checked;
            Record.ShenYan = checkBoxXShenSick.Checked;
            Record.FengShiBing = checkBoxXFengShiSick.Checked;
            Record.DiFangBing = comboBoxExLocalSick.Text;
            Record.QiTaBing = comboBoxExOtherSick.Text;
            Record.ShengBingRiQi = new DateTime(integerInputSickDateYear.Value, integerInputSickDateMonth.Value, integerInputSickDateDay.Value);

            //3.0)青春期
            Record.FloatReserved4 = Record.XingBie ? integerInputYiJing.Value : integerInputChuChao.Value;

            //3)身体机能
            Record.ShenGao = doubleInputHeight.Value.ToFloat();
            Record.TiZhong = doubleInputWeight.Value.ToFloat();
            Record.FloatReserved1 = doubleInputXiongWei.Value.ToFloat();
            Record.FloatReserved2 = doubleInputYaoWei.Value.ToFloat();//腰围
            Record.FloatReserved3 = doubleInputTunWei.Value.ToFloat();//臀围
            Record.FeiHuoLiang = integerInputLungCapacity.Value;
            if(Settings.Default.IsMMHG)
            {
                Record.SuZhangYa = (int)doubleInputSuZhangYa.Value;
                Record.SouSuoYa = (int)doubleInputSouSuoYa.Value;
            }
            else 
            {
                Record.SuZhangYa = (int)(doubleInputSuZhangYa.Value * 7.5);
                Record.SouSuoYa = (int)(doubleInputSouSuoYa.Value * 7.5);
            }
     
            Record.MaiBo = integerInputMaiBo.Value;

            //4)内科
            Record.Xin = comboBoxExXin.Text;
            Record.Fei = comboBoxExFei.Text;
            Record.Gan = comboBoxExGan.Text;
            Record.Pi = comboBoxExPi.Text;

            //5)眼科
            Record.ZuoYan = doubleInputZuoYan.Value.ToFloat();
            Record.YouYan = doubleInputYouYan.Value.ToFloat();
            Record.ShaYan = comboBoxExShaYan.Text;
            Record.JieMoYan = comboBoxExJieMoYan.Text;
            Record.SeJue = comboBoxExSeJue.Text;

            //6)口腔科
            Record.IntReserved1 = integerInputAZuoShang.Value;
            Record.IntReserved2 = integerInputAZuoXia.Value;
            Record.IntReserved3 = integerInputAYouShang.Value;
            Record.IntReserved4 = integerInputAYouXia.Value;


            Record.BZuoShang = integerInputBZuoShang.Value.ToByte();
            Record.BZuoXia = integerInputBZuoXia.Value.ToByte();
            Record.BYouShang = integerInputBYouShang.Value.ToByte();
            Record.BYouXia = integerInputBYouXia.Value.ToByte();

            Record.DZuoShang = integerInputDZuoShang.Value.ToByte();
            Record.DZuoXia = integerInputDZuoXia.Value.ToByte();
            Record.DYouShang = integerInputDYouShang.Value.ToByte();
            Record.DYouXia = integerInputDYouXia.Value.ToByte();

            Record.FZuoShang = integerInputFZuoShang.Value.ToByte();
            Record.FZuoXia = integerInputFZuoXia.Value.ToByte();
            Record.FYouShang = integerInputFYouShang.Value.ToByte();
            Record.FYouXia = integerInputFYouXia.Value.ToByte();

            Record.YaZhou = comboBoxExYaZhou.Text;

            //7)耳鼻咽喉科
            Record.Er = comboBoxExEr.Text;
            Record.Bi = comboBoxExBi.Text;
            Record.BianTaoTi = comboBoxExBianTaoTi.Text;

            //8)外科
            Record.Tou = comboBoxExTou.Text;
            Record.Jing = comboBoxExJing.Text;
            Record.Xiong = comboBoxExXiong.Text;
            Record.JiZhu = comboBoxExJiZhu.Text;
            Record.SiZhi = comboBoxExSiZhi.Text;
            Record.PiFu = comboBoxExPiFu.Text;
            Record.LinBaJie = comboBoxExLinBaJie.Text;

            //9)血型
            Record.XueXing = comboBoxExXueXing.Text;
            Record.XueHongDanBai = doubleInputXueHongDanBai.Value.ToFloat();
            Record.HuiChongLuan = comboBoxExHuiChongLuan.Text;

            //10)肺结核
            Record.KaBa = comboBoxExKaBa.SelectedIndex == 0;
            Record.JieHeJunSu = comboBoxExJieHeJun.Text;

            //11)肝功能
            Record.ZhuanAnMei = comboBoxExZhuanAnMei.Text;
            Record.DanHongSu = comboBoxExDanHongSu.Text;



            //update default value
            Settings.Default.CheckDate = Record.TiJianRiQi;
            Settings.Default.Checker = Record.TiJianDanWei;
            Settings.Default.DefaultGrade = Record.NianJi;
            Settings.Default.DefaultSchool = Record.XueXiao;
            Settings.Default.DefaultClass = Record.BanJi;
            Settings.Default.Save();
            return true;
        }

        private void DisplayDataModelToUI()
        {
            //1) 基本信息
            comboBoxExXueXiao.Text = Record.XueXiao;
            comboBoxExBanJi.Text = Record.BanJi;
            comboBoxExNianJi.Text = Record.NianJi;
       
        textBoxXXingMing.Text =      Record.XingMing ;
        textBoxXXueHao.Text = Record.StringReserved3;
           comboBoxExXingBie.SelectedIndex =  Record.XingBie ?0 : 1;




           integerInputBirthDateYear.Value = Record.ShengRi.Year;
           integerInputBirthDateMonth.Value = Record.ShengRi.Month;
           integerInputBirthDateDay.Value = Record.ShengRi.Day;
           comboBoxExYuHang.SelectedIndex = Record.YuHangHuJi ? 0 : 1;
           integerInputCheckDateYear.Value = Record.TiJianRiQi.Year;
           integerInputCheckDateMonth.Value = Record.TiJianRiQi.Month;
           integerInputCheckDateDay.Value = Record.TiJianRiQi.Day;



           comboBoxExChecker.Text = Record.TiJianDanWei;

            //2) 既往病史
           checkBoxXGanSick.Checked = Record.GanYan;
            checkBoxXFeiSick.Checked = Record.FeiJieHe  ;
            checkBoxXXinSick.Checked = Record.XinZangBing;
            checkBoxXShenSick.Checked = Record.ShenYan;
            checkBoxXFengShiSick.Checked = Record.FengShiBing;
     comboBoxExLocalSick.Text=       Record.DiFangBing ;
     comboBoxExOtherSick.Text = Record.QiTaBing;

     integerInputSickDateYear.Value = Record.ShengBingRiQi.Year;
     integerInputSickDateMonth.Value = Record.ShengBingRiQi.Month;
     integerInputSickDateDay.Value = Record.ShengBingRiQi.Day;

          //3.0)青春期



     if (comboBoxExXingBie.SelectedIndex == 0)
     {
         if ((int)Record.FloatReserved4 == 0 || (int)Record.FloatReserved4 == (int)integerInputYiJing.MinValue)
         {
             integerInputYiJing.ValueObject = null;// Record.FloatReserved1;// = doubleInputXiongWei.Value.ToFloat();
         }
         else
         {
             integerInputYiJing.Value = (int)Record.FloatReserved4;
         }
     }
     else
     {
         if ((int)Record.FloatReserved4 == 0 || (int)Record.FloatReserved4 == (int)integerInputChuChao.MinValue)
         {
             integerInputChuChao.ValueObject = null;// Record.FloatReserved1;// = doubleInputXiongWei.Value.ToFloat();
         }
         else
         {
             integerInputChuChao.Value = (int)Record.FloatReserved4;
         }
     }         

            //3)身体机能
       doubleInputHeight.Value =     Record.ShenGao;// = doubleInputHeight.Value;//.ToFloat();
      doubleInputWeight.Value =      Record.TiZhong;// = doubleInputWeight.Value.ToFloat();

            if((int)Record.FloatReserved1 == 0 || (int)Record.FloatReserved1 == (int)doubleInputXiongWei.MinValue)
            {
                doubleInputXiongWei.ValueObject = null;// Record.FloatReserved1;// = doubleInputXiongWei.Value.ToFloat();
            }
            else 
            {
                doubleInputXiongWei.Value = Record.FloatReserved1;// = doubleInputXiongWei.Value.ToFloat();
            }


            if ((int)Record.FloatReserved2 == 0 || (int)Record.FloatReserved2 == (int)doubleInputYaoWei.MinValue)
            {
                doubleInputYaoWei.ValueObject = null;// Record.FloatReserved1;// = doubleInputXiongWei.Value.ToFloat();
            }
            else
            {
                doubleInputYaoWei.Value = Record.FloatReserved2;// = doubleInputXiongWei.Value.ToFloat();
            }


            if ((int)Record.FloatReserved3 == 0 || (int)Record.FloatReserved3 == (int)doubleInputTunWei.MinValue)
            {
                doubleInputTunWei.ValueObject = null;// Record.FloatReserved1;// = doubleInputXiongWei.Value.ToFloat();
            }
            else
            {
                doubleInputTunWei.Value = Record.FloatReserved3;// = doubleInputXiongWei.Value.ToFloat();
            }
    
         integerInputLungCapacity.Value =    Record.FeiHuoLiang;// = integerInputLungCapacity.Value;
            if(Settings.Default.IsMMHG)
            {
                doubleInputSuZhangYa.Value = Record.SuZhangYa;// = doubleInputSuZhangYa.Value;
                doubleInputSouSuoYa.Value = Record.SouSuoYa;// = doubleInputSouSuoYa.Value;
            }
            else 
            {
                doubleInputSuZhangYa.Value =( Record.SuZhangYa / 7.5);// = doubleInputSuZhangYa.Value;
                doubleInputSouSuoYa.Value = (Record.SouSuoYa / 7.5);// = doubleInputSouSuoYa.Value;
            }
      
          integerInputMaiBo.Value =   Record.MaiBo;// = integerInputMaiBo.Value;

            //4)内科
            
            //Record.Xin = 
                comboBoxExXin.SelectItem(Record.Xin);
            //Record.Fei = comboBoxExFei.Text;
                comboBoxExFei.SelectItem(Record.Fei);
           // Record.Gan = comboBoxExGan.Text;
                comboBoxExFei.SelectItem(Record.Gan);            
            //Record.Pi = comboBoxExPi.Text;
                comboBoxExPi.SelectItem(Record.Pi);

            //5)眼科
            //Record.ZuoYan = doubleInputZuoYan.Value.ToFloat();
            doubleInputZuoYan.Value = Record.ZuoYan;
           // Record.YouYan = doubleInputYouYan.Value.ToFloat();
            doubleInputYouYan.Value = Record.YouYan;
            //Record.ShaYan = comboBoxExShaYan.Text;
            comboBoxExShaYan.SelectItem(Record.ShaYan);
            //Record.JieMoYan = comboBoxExJieMoYan.Text;
            comboBoxExJieMoYan.SelectItem(Record.JieMoYan);
            //Record.SeJue = comboBoxExSeJue.Text;
            comboBoxExSeJue.SelectItem(Record.SeJue);

            //6)口腔科
            integerInputAZuoShang.Value = Record.IntReserved1;
            integerInputAZuoXia.Value = Record.IntReserved2;
            integerInputAYouShang.Value = Record.IntReserved3;
            integerInputAYouXia.Value = Record.IntReserved4;


            //Record.BZuoShang = integerInputBZuoShang.Value.ToByte();
            integerInputBZuoShang.Value = Record.BZuoShang;
            //Record.BZuoXia = integerInputBZuoXia.Value.ToByte();
            integerInputBZuoXia.Value = Record.BZuoXia;
            //Record.BYouShang = integerInputBYouShang.Value.ToByte()
            integerInputBYouShang.Value = Record.BYouShang;
            //Record.BYouXia = integerInputBYouXia.Value.ToByte();
            integerInputBYouXia.Value = Record.BYouXia;

            //Record.DZuoShang = integerInputDZuoShang.Value.ToByte();
            integerInputDZuoShang.Value = Record.DZuoShang;
           // Record.DZuoXia = integerInputDZuoXia.Value.ToByte();
            integerInputDZuoXia.Value = Record.DZuoXia;
           // Record.DYouShang = integerInputDYouShang.Value.ToByte();
            integerInputDYouShang.Value = Record.DYouShang;
           // Record.DYouXia = integerInputDYouXia.Value.ToByte();
            integerInputDYouXia.Value = Record.DYouXia;

            //Record.FZuoShang = integerInputFZuoShang.Value.ToByte();
            integerInputFZuoShang.Value = Record.FZuoShang;
            //Record.FZuoXia = integerInputFZuoXia.Value.ToByte();
            integerInputFZuoXia.Value = Record.FZuoXia;
            //Record.FYouShang = integerInputFYouShang.Value.ToByte();
            integerInputFYouShang.Value = Record.FYouShang;
            //Record.FYouXia = integerInputFYouXia.Value.ToByte();
            integerInputFYouXia.Value = Record.FYouXia;

            //Record.YaZhou = comboBoxExYaZhou.Text;
            comboBoxExYaZhou.SelectItem(Record.YaZhou);

            //7)耳鼻咽喉科
            //Record.Er = comboBoxExEr.Text;
            comboBoxExEr.SelectItem(Record.Er);
            //Record.Bi = comboBoxExBi.Text;
            comboBoxExBi.SelectItem(Record.Bi);
            //Record.BianTaoTi = comboBoxExBianTaoTi.Text;
            comboBoxExBianTaoTi.SelectItem(Record.BianTaoTi);

            //8)外科
            //Record.Tou = comboBoxExTou.Text;
            comboBoxExTou.SelectItem(Record.Tou);
            //Record.Jing = comboBoxExJing.Text;
            comboBoxExJing.SelectItem(Record.Jing);
            //Record.Xiong = comboBoxExXiong.Text;
            comboBoxExXiong.SelectItem(Record.Xiong);
            //Record.JiZhu = comboBoxExJiZhu.Text;
            comboBoxExJiZhu.SelectItem(Record.JiZhu);
            //Record.SiZhi = comboBoxExSiZhi.Text;
            comboBoxExSiZhi.SelectItem(Record.SiZhi);
            //Record.PiFu = comboBoxExPiFu.Text;
            comboBoxExPiFu.SelectItem(Record.PiFu);
            //Record.LinBaJie = comboBoxExLinBaJie.Text;
            comboBoxExLinBaJie.SelectItem(Record.LinBaJie);

            //9)血型
            //Record.XueXing = comboBoxExXueXing.Text;
            comboBoxExXueXing.SelectItem(Record.XueXing);
            //Record.XueHongDanBai = doubleInputXueHongDanBai.Value.ToFloat();
            doubleInputXueHongDanBai.Value = Record.XueHongDanBai;
            //Record.HuiChongLuan = comboBoxExHuiChongLuan.Text;
            comboBoxExHuiChongLuan.SelectItem(Record.HuiChongLuan);

            //10)肺结核
         //   Record.KaBa = comboBoxExKaBa.SelectedIndex == 0;
            comboBoxExKaBa.SelectedIndex = Record.KaBa ? 0 : 1;
   //         Record.JieHeJunSu = comboBoxExJieHeJun.Text;
            comboBoxExJieHeJun.SelectItem(Record.JieHeJunSu);
            //11)肝功能
       //     Record.ZhuanAnMei = comboBoxExZhuanAnMei.Text;
            comboBoxExZhuanAnMei.SelectItem(Record.ZhuanAnMei);
            //Record.DanHongSu = comboBoxExDanHongSu.Text;
            comboBoxExDanHongSu.SelectItem(Record.DanHongSu);
        }

        internal void UpdateRecod(int id)
        {
           var r = DBHelper.GetRecord(id);
            if(r != null)
            {
                Record = r;
                DisplayDataModelToUI();
            }
        }

        internal void SwitchShortCut(bool enable)
        {
            if(enable)
            {
                if (IsUpdateMode)
                {
                    btnSaveAndNewEx.Shortcuts.Clear();
                    btnSaveAndNewEx.Shortcuts.Add(eShortcut.F4);
                }
                else
                {
                    btnSaveAndNewEx.Shortcuts.Clear();        
                    btnSaveAndNewEx.Shortcuts.Add(eShortcut.F4);
                    btnSaveEx.Shortcuts.Clear();
                    btnSaveEx.Shortcuts.Add(eShortcut.F3);
                    btnNew.Shortcuts.Clear();
                    btnNew.Shortcuts.Add(eShortcut.F2);
                }
            }
            else 
            {
                if (IsUpdateMode)
                {
                    btnSaveAndNewEx.Shortcuts.Clear();
                }
                else
                {
                    btnSaveAndNewEx.Shortcuts.Clear();
                    btnSaveEx.Shortcuts.Clear();
                    btnNew.Shortcuts.Clear();
                }
            }
           
        }

        internal void UpdateBloodPressureUnit()
        {
          if(Settings.Default.IsMMHG)
          {
              lblBloodPresureUnit1.Text = "(mmHg)";
            //  lblBloodPresureUnit2.Text = lblBloodPresureUnit1.Text;

              var v = doubleInputSuZhangYa.Value;
              doubleInputSuZhangYa.MinValue = 0;
              doubleInputSuZhangYa.MaxValue = 300;
              doubleInputSuZhangYa.DisplayFormat = "f0";

              if (v != doubleInputSuZhangYa.MinValue)
              {
                  doubleInputSuZhangYa.Value = (int)(v * 7.5);
              }

              v = doubleInputSouSuoYa.Value;
              doubleInputSouSuoYa.MinValue = 0;
              doubleInputSouSuoYa.MaxValue = 300;
              doubleInputSouSuoYa.DisplayFormat = "f0";

              if (v != doubleInputSouSuoYa.MinValue)
              {
                  doubleInputSouSuoYa.Value = (int)(v * 7.5);
              }
             
            
          }
          else 
          {
              lblBloodPresureUnit1.Text = "(kPa)";
            //  lblBloodPresureUnit2.Text = lblBloodPresureUnit1.Text;

              var v = doubleInputSuZhangYa.Value;
              doubleInputSuZhangYa.MinValue = 0;
              doubleInputSuZhangYa.MaxValue = 40;
              doubleInputSuZhangYa.DisplayFormat = "f1";
              if (v != doubleInputSuZhangYa.MinValue)
              {
                  doubleInputSuZhangYa.Value = (int)(v / 7.5);
              }

              v = doubleInputSouSuoYa.Value;
              doubleInputSouSuoYa.MinValue = 0;
              doubleInputSouSuoYa.MaxValue = 40;
              doubleInputSouSuoYa.DisplayFormat = "f1";
              if (v != doubleInputSouSuoYa.MinValue)
              {
                  doubleInputSouSuoYa.Value = (int)(v / 7.5);
              }
          
          }
        }

        private void comboBoxExXingBie_SelectedIndexChanged(object sender, EventArgs e)
        {
            integerInputYiJing.Enabled = comboBoxExXingBie.SelectedIndex == 0;
            integerInputChuChao.Enabled = !integerInputYiJing.Enabled;
        }
    }
}
