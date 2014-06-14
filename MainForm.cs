using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using DevComponents.DotNetBar.Controls;

using SchoolHealth.Properties;
using System.Threading;

namespace SchoolHealth
{
    public partial class MainForm : DevComponents.DotNetBar.Office2007Form
    {

        #region Members
        RecordInputControl newControl;
        RecordInputControl updateControl;
        //  AutoResetEvent printOneReportEvent;
        //  AutoResetEvent /*printResetEvent*/;
        // bool isPrinting = false;
        ExportType expType = ExportType.Word;
        Microsoft.Reporting.WinForms.ReportViewer curViewer = null;
        #endregion

        #region Ctor
        public MainForm()
        {
            InitializeComponent();
            MessageBoxEx.UseSystemLocalizedString = true;
        }
        #endregion

        #region Overide Methods

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);

            Utility.SaveInvoke(this, () =>
                {
                    //flag == 1, ver 1002
                    //flag == 2, ver 1003
                    if (Settings.Default.UpdateFlag != 2)
                    {
#if DEBUG1
                        Utility.ShowInformation("需要更新");
                        Settings.Default.UpdateFlag = 2;
                        Settings.Default.Save();
#else              
                        DBHelper.UpdateStatisticResult(true);             
#endif

                    }
                    else
                    {
#if DEBUG
                        Utility.ShowInformation("不需要更新");
#endif
                    }
                }, true, false);

        }

        protected override void OnClosing(CancelEventArgs e)
        {
            if (Settings.Default.AutoBackup)
            {
                DBHelper.AutoBackup();
            }

            base.OnClosing(e);


        }


        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                SendKeys.SendWait("{Tab}");
                return true;
            }

            if (keyData == (Keys.Enter | Keys.Shift))
            {
                SendKeys.SendWait("+{Tab}");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #endregion

        #region Helper Methods
        private void SearchRecord(bool byProgressFrom, string title)
        {
            superTabControlHistory.SelectedTabIndex = 0;

            if (byProgressFrom)
            {
                ProgressForm.QueueWorkingItem(() =>
                {
                    BindingSource bs = new BindingSource();
                    string strSql = GetSearchSQL();

                    var dataSet = DBHelper.SchoolHealthQueryDataSet(strSql);
                    bs.DataSource = dataSet.Tables[0].DefaultView;
                    this.bindingNavigator.BindingSource = bs;
                    this.dataGridViewXHistory.DataSource = bs;

                }, title, this);
            }
            else
            {
                BindingSource bs = new BindingSource();
                string strSql = GetSearchSQL();

                var dataSet = DBHelper.SchoolHealthQueryDataSet(strSql);
                bs.DataSource = dataSet.Tables[0].DefaultView;
                this.bindingNavigator.BindingSource = bs;
                this.dataGridViewXHistory.DataSource = bs;
            }
        }

        internal void Initialize(StartUpHelper startUpHelper)
        {
            // startUpHelper.UpdateMsg("正在初始化数据源");
            try
            {
                Thread.Sleep(500);


                //    Utility.ShowInformation("1");
                newControl = new RecordInputControl();
                newControl.Dock = DockStyle.Fill;
                newControl.Initialize();
                newControl.SwitchShortCut(true);
                superTabControlPanelNew.Controls.Add(newControl);
                Thread.Sleep(500);
                //     Utility.ShowInformation("2");
                updateControl = new RecordInputControl();
                updateControl.Dock = DockStyle.Fill;

                updateControl.Initialize();
                updateControl.SwitchShortCut(false);
                updateControl.IsUpdateMode = true;
                superTabControlPanelUpdate.Controls.Add(updateControl);
                Thread.Sleep(500);

                //     Utility.ShowInformation("3");

                //打印机
                comboBoxExPrint.SelectedIndexChanged -= comboBoxExPrint_SelectedIndexChanged;
                comboBoxExPrint.DataSource = Utility.LoadPrinters();
                var dp = Properties.Settings.Default.DefaultPrinter;
                if (!dp.IsNullOrEmpty())
                {
                    var index = comboBoxExPrint.FindStringExact(dp);
                    if (index > -1)
                    {
                        comboBoxExPrint.SelectedIndex = index;
                        comboBoxExPrint_SelectedIndexChanged(comboBoxExPrint, null);
                    }
                }
                else if (comboBoxExPrint.Items.Count > 0)
                {
                    comboBoxExPrint_SelectedIndexChanged(comboBoxExPrint, null);
                }
                comboBoxExPrint.SelectedIndexChanged += comboBoxExPrint_SelectedIndexChanged;
                Thread.Sleep(500);

                //        Utility.ShowInformation("4");
                //UI Style
                comboBoxExUIStyle.SelectedIndexChanged -= comboBoxExUIStyle_SelectedIndexChanged;
                comboBoxExUIStyle.DataSource = Enum.GetNames(typeof(eStyle));
                var ds = Properties.Settings.Default.DefaultStyle;
                if (!ds.IsNullOrEmpty())
                {
                    var index = comboBoxExUIStyle.FindStringExact(ds);
                    if (index > -1)
                    {
                        comboBoxExUIStyle.SelectedIndex = index;
                        comboBoxExUIStyle_SelectedIndexChanged(comboBoxExUIStyle, null);
                    }
                }
                else if (comboBoxExUIStyle.Items.Count > 0)
                {
                    comboBoxExUIStyle_SelectedIndexChanged(comboBoxExUIStyle, null);
                }
                comboBoxExUIStyle.SelectedIndexChanged += comboBoxExUIStyle_SelectedIndexChanged;


                LoadSetting();
                Settings.Default.PropertyChanged += new PropertyChangedEventHandler(Default_PropertyChanged);
            }


            catch (Exception ex)
            {
                Utility.ShowInformation(ex.Message);
            }
            finally
            {
                startUpHelper.SetCompleted();
            }





        }

        private void LoadSetting()
        {
            checkBoxMMHG.Checked = Settings.Default.IsMMHG;
            checkBoxXKPA.Checked = !Settings.Default.IsMMHG;
        }

        void Default_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "IsMMHG")
            {
                updateControl.UpdateBloodPressureUnit();
                newControl.UpdateBloodPressureUnit();
            }
        }

        private void RefreshSchool()
        {
            textBoxDropDownSchool.DropDownItems.Clear();
            string grade = textBoxDropDownSchool.Text;

            var ret = DBHelper.ListSchool();

            if (ret.Count > 0)
            {
                textBoxDropDownSchool.AutoCompleteCustomSource.AddRange(ret.ToArray());
                foreach (var s in ret)
                {
                    ButtonItem bi = new ButtonItem("", s);
                    bi.Click += delegate
                    {
                        textBoxDropDownSchool.Text = bi.Text;
                        RefreshGrade();
                        RefreshClass();
                    };
                    textBoxDropDownSchool.DropDownItems.Add(bi);
                }
            }
        }

        private void RefreshGrade()
        {
            textBoxDropDownGrade.DropDownItems.Clear();
            string school = textBoxDropDownSchool.Text;

            var ret = DBHelper.ListGrade(school);

            if (ret.Count > 0)
            {

                var defaultGrades = ConfigurationHelper.DataSourceByKey(ConfigurationHelper.年级);

                var tmp = new List<string>();

                foreach (var g in defaultGrades)
                {
                    if (ret.Contains(g))
                    {
                        tmp.Add(g);
                    }
                }


                textBoxDropDownGrade.AutoCompleteCustomSource.AddRange(tmp.ToArray());
                foreach (var s in tmp)
                {
                    ButtonItem bi = new ButtonItem("", s);
                    bi.Click += delegate
                    {
                        textBoxDropDownGrade.Text = bi.Text;
                        RefreshClass();
                    };
                    textBoxDropDownGrade.DropDownItems.Add(bi);
                }
            }
        }

        private void RefreshClass()
        {
            textBoxDropDownClass.DropDownItems.Clear();
            string grade = textBoxDropDownGrade.Text;

            var ret = DBHelper.ListClass(textBoxDropDownSchool.Text, grade);

            if (ret.Count > 0)
            {
                textBoxDropDownClass.AutoCompleteCustomSource.AddRange(ret.ToArray());
                foreach (var s in ret)
                {
                    ButtonItem bi = new ButtonItem("", s);
                    bi.Click += delegate
                    {
                        textBoxDropDownClass.Text = bi.Text;
                    };
                    textBoxDropDownClass.DropDownItems.Add(bi);
                }
            }
        }

        //private void RefreshHistoryData()
        //{
        //    ProgressForm.QueueWorkingItem(() =>
        //        {
        //            try
        //            {
        //                if (this.bindingNavigator.BindingSource == null)
        //                {
        //                    BindingSource bs = new BindingSource();
        //                    var dataSet = DBHelper.GetAll();
        //                    bs.DataSource = dataSet.Tables[0].DefaultView;
        //                    this.bindingNavigator.BindingSource = bs;
        //                    this.bindingNavigator.AddNewRecordButton.Visible = false;
        //                    this.bindingNavigator.DeleteButton.Visible = false;
        //                    this.dataGridViewXHistory.AutoGenerateColumns = false;
        //                    this.dataGridViewXHistory.Columns.AddRange(DBHelper.DefaultColumns.ToArray());
        //                    this.dataGridViewXHistory.DataSource = bs;
        //                //    isNewRecordAddedButNotViewHistory = false;
        //                    RefreshSchool();
        //                    RefreshClass();
        //                }

        //                if (this.bindingNavigator.BindingSource != null && isNewRecordAddedButNotViewHistory)
        //                {
        //                    RefreshSchool();
        //                    RefreshClass();
        //                    SearchRecord(false, string.Empty);
        //            //        isNewRecordAddedButNotViewHistory = false;
        //                }
        //            }
        //            catch
        //            {

        //            }
        //            finally
        //            {

        //            }
        //        }, "正在刷新数据...", this);


        //}

        //private void ResetComboBoxSelection()
        //{

        //    foreach (var gp in this.expandPanels)
        //    {
        //        foreach (Control child in gp.Controls)
        //        {
        //            ComboBoxEx cbx = child as ComboBoxEx;
        //            if (cbx != null && cbx.DropDownStyle == ComboBoxStyle.DropDownList)
        //            {
        //                cbx.SelectedIndex = cbx.Items.Count > 0 ? 0 : -1;
        //            }
        //        }
        //    }
        //}

        //private void InitializeComboBoxSource()
        //{
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.年级, comboBoxExNianJi);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.性别, comboBoxExXingBie);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.余杭户籍, comboBoxExYuHang);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.心, comboBoxExXin);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.肺, comboBoxExFei);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.肝, comboBoxExGan);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.脾, comboBoxExPi);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.沙眼, comboBoxExShaYan);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.结膜炎, comboBoxExJieMoYan);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.色觉, comboBoxExSeJue);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.牙周, comboBoxExYaZhou);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.耳, comboBoxExEr);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.鼻, comboBoxExBi);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.扁桃体, comboBoxExBianTaoTi);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.头部, comboBoxExTou);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.颈部, comboBoxExJing);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.胸部, comboBoxExXiong);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.脊柱, comboBoxExJiZhu);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.四肢, comboBoxExSiZhi);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.皮肤, comboBoxExPiFu);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.淋巴结, comboBoxExLinBaJie);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.血型, comboBoxExXueXing);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.蛔虫卵, comboBoxExHuiChongLuan);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.卡疤, comboBoxExKaBa);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.结核菌素试验, comboBoxExJieHeJun);

        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.谷丙转氨酶, comboBoxExZhuanAnMei);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.胆红素, comboBoxExDanHongSu);

        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.年级, textBoxDropDownGrade, RefreshClass);
        //    // ConfigurationHelper.BindDataSource(ConfigurationHelper.Excel数据类型, comboBoxExExcelExportType);


        //    //打印机
        //    //comboBoxExPrint.SelectedIndexChanged -= comboBoxExPrint_SelectedIndexChanged;
        //    //comboBoxExPrint.DataSource = Utility.LoadPrinters();
        //    //var dp = Properties.Settings.Default.DefaultPrinter;
        //    //if (!dp.IsNullOrEmpty())
        //    //{
        //    //    var index = comboBoxExPrint.FindStringExact(dp);
        //    //    if (index > -1)
        //    //    {
        //    //        comboBoxExPrint.SelectedIndex = index;
        //    //        comboBoxExPrint_SelectedIndexChanged(comboBoxExPrint, null);
        //    //    }
        //    //}
        //    //else if (comboBoxExPrint.Items.Count > 0)
        //    //{
        //    //    comboBoxExPrint_SelectedIndexChanged(comboBoxExPrint, null);
        //    //}
        //    //comboBoxExPrint.SelectedIndexChanged += comboBoxExPrint_SelectedIndexChanged;

        //    //UI Style
        //    comboBoxExUIStyle.SelectedIndexChanged -= comboBoxExUIStyle_SelectedIndexChanged;
        //    comboBoxExUIStyle.DataSource = Enum.GetNames(typeof(eStyle));
        //    var ds = Properties.Settings.Default.DefaultStyle;
        //    if (!ds.IsNullOrEmpty())
        //    {
        //        var index = comboBoxExUIStyle.FindStringExact(ds);
        //        if (index > -1)
        //        {
        //            comboBoxExUIStyle.SelectedIndex = index;
        //            comboBoxExUIStyle_SelectedIndexChanged(comboBoxExUIStyle, null);
        //        }
        //    }
        //    else if (comboBoxExUIStyle.Items.Count > 0)
        //    {
        //        comboBoxExUIStyle_SelectedIndexChanged(comboBoxExUIStyle, null);
        //    }
        //    comboBoxExUIStyle.SelectedIndexChanged += comboBoxExUIStyle_SelectedIndexChanged;

        //}

        //private void InitInputValue()
        //{
        //    integerInputBYouShang.Value = 0;
        //    integerInputBYouXia.Value = 0;
        //    integerInputBZuoShang.Value = 0;
        //    integerInputBZuoXia.Value = 0;

        //    integerInputDYouShang.Value = 0;
        //    integerInputDYouXia.Value = 0;
        //    integerInputDZuoShang.Value = 0;
        //    integerInputDZuoXia.Value = 0;

        //    integerInputFYouShang.Value = 0;
        //    integerInputFYouXia.Value = 0;
        //    integerInputFZuoShang.Value = 0;
        //    integerInputFZuoXia.Value = 0;

        //    doubleInputHeight.Value = 0;
        //    doubleInputWeight.Value = 0;
        //    integerInputLungCapacity.Value = 0;
        //    integerInputShuZhangYa.Value = 0;
        //    integerInputSouSuoYa.Value = 0;
        //    integerInputMaiBo.Value = 0;

        //    doubleInputXueHongDanBai.Value = 0;

        //    doubleInputHeight.ValueObject = null;
        //    doubleInputWeight.ValueObject = null;
        //    integerInputLungCapacity.ValueObject = null;
        //    integerInputShuZhangYa.ValueObject = null;
        //    integerInputSouSuoYa.ValueObject = null;
        //    integerInputMaiBo.ValueObject = null;

        //    doubleInputXueHongDanBai.ValueObject = null;





        //    textBoxXXingMing.Text = string.Empty;
        //    doubleInputZuoYan.Value = 5.0;
        //    doubleInputYouYan.Value = 5.0;


        //    integerInputBirthDateMonth.ValueObject = null;
        //    integerInputBirthDateDay.ValueObject = null;

        //    integerInputSickDateYear.ValueObject = null;
        //    integerInputSickDateMonth.ValueObject = null;
        //    integerInputSickDateDay.ValueObject = null;
        //}

        //private void ScrollToVisible(Control ctrl)
        //{
        //    groupPanelRecord.ScrollControlIntoView(ctrl);
        //    ctrl.Focus();
        //}

        //private bool UpdateUIInputToDataModel()
        //{


        //    //check

        //    if (comboBoxExXueXiao.Text.IsNullOrEmpty())
        //    {
        //        Utility.ShowInformation("学校不可为空!", this);
        //        ScrollToVisible(comboBoxExXueXiao);
        //        return false;
        //    }

        //    if (comboBoxExBanJi.Text.IsNullOrEmpty())
        //    {
        //        Utility.ShowInformation("班级不可为空!", this);
        //        ScrollToVisible(comboBoxExBanJi);
        //        return false;
        //    }

        //    if (textBoxXXingMing.Text.IsNullOrEmpty())
        //    {
        //        Utility.ShowInformation("姓名不可为空!", this);
        //        ScrollToVisible(textBoxXXingMing);
        //        return false;
        //    }

        //    if (comboBoxExChecker.Text.IsNullOrEmpty())
        //    {
        //        Utility.ShowInformation("体检单位不可为空!", this);
        //        ScrollToVisible(comboBoxExChecker);
        //        return false;
        //    }

        //    if ((int)doubleInputHeight.Value == 0 ||  (int)doubleInputHeight.Value == (int)doubleInputHeight.MinValue)
        //    {
        //        Utility.ShowInformation("请输入身高!", this);
        //        ScrollToVisible(doubleInputHeight);
        //        return false;
        //    }

        //    if ((int)doubleInputWeight.Value == 0 || (int)doubleInputWeight.Value == (int)doubleInputWeight.MinValue)
        //    {
        //        Utility.ShowInformation("请输入体重!", this);
        //        ScrollToVisible(doubleInputWeight);
        //        return false;
        //    }

        //    //if (integerInputLungCapacity.Value == 0)
        //    //{
        //    //    Utility.ShowInformation("请输入肺活量!", this);
        //    //    ScrollToVisible(integerInputLungCapacity);
        //    //    return false;
        //    //}

        //    if (integerInputShuZhangYa.Value == 0 || integerInputShuZhangYa.Value == integerInputShuZhangYa.MinValue)
        //    {
        //        Utility.ShowInformation("请输入舒张压!", this);
        //        ScrollToVisible(integerInputShuZhangYa);
        //        return false;
        //    }

        //    if (integerInputSouSuoYa.Value == 0 || integerInputSouSuoYa.Value == integerInputSouSuoYa.MinValue)
        //    {
        //        Utility.ShowInformation("请输入收缩压!", this);
        //        ScrollToVisible(integerInputSouSuoYa);
        //        return false;
        //    }

        //    if (integerInputMaiBo.Value == 0 || integerInputMaiBo.Value == integerInputMaiBo.MinValue)
        //    {
        //        Utility.ShowInformation("请输入脉搏!", this);
        //        ScrollToVisible(integerInputMaiBo);
        //        return false;
        //    }


        //    //if ((int)doubleInputXueHongDanBai.Value == 0)
        //    //{
        //    //    Utility.ShowInformation("请输入血红蛋白值!", this);
        //    //    ScrollToVisible(doubleInputXueHongDanBai);
        //    //    return false;
        //    //}


        //    //assign 
        //    //1) 基本信息
        //    Record.XueXiao = comboBoxExXueXiao.Text;
        //    Record.BanJi = comboBoxExBanJi.Text;
        //    Record.NianJi = comboBoxExNianJi.Text;
        //    Record.XingMing = textBoxXXingMing.Text;
        //    Record.XingBie = comboBoxExXingBie.SelectedIndex == 0;
        //    Record.ShengRi = new DateTime(integerInputBirthDateYear.Value, integerInputBirthDateMonth.Value, integerInputBirthDateDay.Value);
        //    Record.YuHangHuJi = comboBoxExYuHang.SelectedIndex == 0;
        //    Record.TiJianRiQi = new DateTime(integerInputCheckDateYear.Value, integerInputCheckDateMonth.Value, integerInputCheckDateDay.Value);
        //    Record.TiJianDanWei = comboBoxExChecker.Text;

        //    //2) 既往病史
        //    Record.GanYan = checkBoxXGanSick.Checked;
        //    Record.FeiJieHe = checkBoxXFeiSick.Checked;
        //    Record.XinZangBing = checkBoxXXinSick.Checked;
        //    Record.ShenYan = checkBoxXShenSick.Checked;
        //    Record.FengShiBing = checkBoxXFengShiSick.Checked;
        //    Record.DiFangBing = comboBoxExLocalSick.Text;
        //    Record.QiTaBing = comboBoxExOtherSick.Text;
        //    Record.ShengBingRiQi = new DateTime(integerInputSickDateYear.Value, integerInputSickDateMonth.Value, integerInputSickDateDay.Value);

        //    //3)身体机能
        //    Record.ShenGao = doubleInputHeight.Value.ToFloat();
        //    Record.TiZhong = doubleInputWeight.Value.ToFloat();
        //    Record.FeiHuoLiang = integerInputLungCapacity.Value;
        //    Record.SuZhangYa = integerInputShuZhangYa.Value;
        //    Record.SouSuoYa = integerInputSouSuoYa.Value;
        //    Record.MaiBo = integerInputMaiBo.Value;

        //    //4)内科
        //    Record.Xin = comboBoxExXin.Text;
        //    Record.Fei = comboBoxExFei.Text;
        //    Record.Gan = comboBoxExGan.Text;
        //    Record.Pi = comboBoxExPi.Text;

        //    //5)眼科
        //    Record.ZuoYan = doubleInputZuoYan.Value.ToFloat();
        //    Record.YouYan = doubleInputYouYan.Value.ToFloat();
        //    Record.ShaYan = comboBoxExShaYan.Text;
        //    Record.JieMoYan = comboBoxExJieMoYan.Text;
        //    Record.SeJue = comboBoxExSeJue.Text;

        //    //6)口腔科
        //    Record.BZuoShang = integerInputBZuoShang.Value.ToByte();
        //    Record.BZuoXia = integerInputBZuoXia.Value.ToByte();
        //    Record.BYouShang = integerInputBYouShang.Value.ToByte();
        //    Record.BYouXia = integerInputBYouXia.Value.ToByte();

        //    Record.DZuoShang = integerInputDZuoShang.Value.ToByte();
        //    Record.DZuoXia = integerInputDZuoXia.Value.ToByte();
        //    Record.DYouShang = integerInputDYouShang.Value.ToByte();
        //    Record.DYouXia = integerInputDYouXia.Value.ToByte();

        //    Record.FZuoShang = integerInputFZuoShang.Value.ToByte();
        //    Record.FZuoXia = integerInputFZuoXia.Value.ToByte();
        //    Record.FYouShang = integerInputFYouShang.Value.ToByte();
        //    Record.FYouXia = integerInputFYouXia.Value.ToByte();

        //    Record.YaZhou = comboBoxExYaZhou.Text;

        //    //7)耳鼻咽喉科
        //    Record.Er = comboBoxExEr.Text;
        //    Record.Bi = comboBoxExBi.Text;
        //    Record.BianTaoTi = comboBoxExBianTaoTi.Text;

        //    //8)外科
        //    Record.Tou = comboBoxExTou.Text;
        //    Record.Jing = comboBoxExJing.Text;
        //    Record.Xiong = comboBoxExXiong.Text;
        //    Record.JiZhu = comboBoxExJiZhu.Text;
        //    Record.SiZhi = comboBoxExSiZhi.Text;
        //    Record.PiFu = comboBoxExPiFu.Text;
        //    Record.LinBaJie = comboBoxExLinBaJie.Text;

        //    //9)血型
        //    Record.XueXing = comboBoxExXueXing.Text;
        //    Record.XueHongDanBai = doubleInputXueHongDanBai.Value.ToFloat();
        //    Record.HuiChongLuan = comboBoxExHuiChongLuan.Text;

        //    //10)肺结核
        //    Record.KaBa = comboBoxExKaBa.SelectedIndex == 0;
        //    Record.JieHeJunSu = comboBoxExJieHeJun.Text;

        //    //11)肝功能
        //    Record.ZhuanAnMei = comboBoxExZhuanAnMei.Text;
        //    Record.DanHongSu = comboBoxExDanHongSu.Text;



        //    //update default value
        //    Settings.Default.CheckDate = Record.TiJianRiQi;
        //    Settings.Default.Checker = Record.TiJianDanWei;
        //    Settings.Default.DefaultGrade = Record.NianJi;
        //    Settings.Default.DefaultSchool = Record.XueXiao;
        //    Settings.Default.DefaultClass = Record.BanJi;
        //    Settings.Default.Save();
        //    return true;
        //}

        //private void DisplayDataModelToUI()
        //{
        //    //1) 基本信息
        //    comboBoxExXueXiao.Text = Record.XueXiao;
        //    comboBoxExBanJi.Text = Record.BanJi;
        //    comboBoxExNianJi.Text = Record.NianJi;
        //    //Record.XingMing = textBoxXXingMing.Text;
        //    //Record.XingBie comboBoxExXingBie.SelectedIndex == 0;
        //    //Record.ShengRi = new DateTime(integerInputBirthDateYear.Value, integerInputBirthDateMonth.Value, integerInputBirthDateDay.Value);
        //    //Record.YuHangHuJi comboBoxExYuHang.SelectedIndex == 0;
        //    //Record.TiJianRiQi = new DateTime(integerInputCheckDateYear.Value, integerInputCheckDateMonth.Value, integerInputCheckDateDay.Value);
        //    //Record.TiJianDanWei comboBoxExChecker.Text;

        //    ////2) 既往病史
        //    //Record.GanYan = checkBoxXGanSick.Checked;
        //    //Record.FeiJieHe = checkBoxXFeiSick.Checked;
        //    //Record.XinZangBing = checkBoxXXinSick.Checked;
        //    //Record.ShenYan = checkBoxXShenSick.Checked;
        //    //Record.FengShiBing = checkBoxXFengShiSick.Checked;
        //    //Record.DiFangBing comboBoxExLocalSick.Text;
        //    //Record.QiTaBing comboBoxExOtherSick.Text;
        //    //Record.ShengBingRiQi = new DateTime(integerInputSickDateYear.Value, integerInputSickDateMonth.Value, integerInputSickDateDay.Value);

        //    ////3)身体机能
        //    //Record.ShenGao = doubleInputHeight.Value.ToFloat();
        //    //Record.TiZhong = doubleInputWeight.Value.ToFloat();
        //    //Record.FeiHuoLiang = integerInputLungCapacity.Value;
        //    //Record.SuZhangYa = integerInputShuZhangYa.Value;
        //    //Record.SouSuoYa = integerInputSouSuoYa.Value;
        //    //Record.MaiBo = integerInputMaiBo.Value;

        //    ////4)内科
        //    //Record.Xin comboBoxExXin.Text;
        //    //Record.Fei comboBoxExFei.Text;
        //    //Record.Gan comboBoxExGan.Text;
        //    //Record.Pi comboBoxExPi.Text;

        //    ////5)眼科
        //    //Record.ZuoYan = doubleInputZuoYan.Value.ToFloat();
        //    //Record.YouYan = doubleInputYouYan.Value.ToFloat();
        //    //Record.ShaYan comboBoxExShaYan.Text;
        //    //Record.JieMoYan comboBoxExJieMoYan.Text;
        //    //Record.SeJue comboBoxExSeJue.Text;

        //    ////6)口腔科
        //    //Record.BZuoShang = integerInputBZuoShang.Value.ToByte();
        //    //Record.BZuoXia = integerInputBZuoXia.Value.ToByte();
        //    //Record.BYouShang = integerInputBYouShang.Value.ToByte();
        //    //Record.BYouXia = integerInputBYouXia.Value.ToByte();

        //    //Record.DZuoShang = integerInputDZuoShang.Value.ToByte();
        //    //Record.DZuoXia = integerInputDZuoXia.Value.ToByte();
        //    //Record.DYouShang = integerInputDYouShang.Value.ToByte();
        //    //Record.DYouXia = integerInputDYouXia.Value.ToByte();

        //    //Record.FZuoShang = integerInputFZuoShang.Value.ToByte();
        //    //Record.FZuoXia = integerInputFZuoXia.Value.ToByte();
        //    //Record.FYouShang = integerInputFYouShang.Value.ToByte();
        //    //Record.FYouXia = integerInputFYouXia.Value.ToByte();

        //    //Record.YaZhou comboBoxExYaZhou.Text;

        //    ////7)耳鼻咽喉科
        //    //Record.Er comboBoxExEr.Text;
        //    //Record.Bi comboBoxExBi.Text;
        //    //Record.BianTaoTi comboBoxExBianTaoTi.Text;

        //    ////8)外科
        //    //Record.Tou comboBoxExTou.Text;
        //    //Record.Jing comboBoxExJing.Text;
        //    //Record.Xiong comboBoxExXiong.Text;
        //    //Record.JiZhu comboBoxExJiZhu.Text;
        //    //Record.SiZhi comboBoxExSiZhi.Text;
        //    //Record.PiFu comboBoxExPiFu.Text;
        //    //Record.LinBaJie comboBoxExLinBaJie.Text;

        //    ////9)血型
        //    //Record.XueXing comboBoxExXueXing.Text;
        //    //Record.XueHongDanBai = doubleInputXueHongDanBai.Value.ToFloat();
        //    //Record.HuiChongLuan comboBoxExHuiChongLuan.Text;

        //    ////10)肺结核
        //    //Record.KaBa comboBoxExKaBa.SelectedIndex == 0;
        //    //Record.JieHeJunShiYan comboBoxExJieHeJun.Text;

        //    ////11)肝功能
        //    //Record.ZhuanAnMei comboBoxExZhuanAnMei.Text;
        //    //Record.DanHongSu comboBoxExDanHongSu.Text;
        //}

        private string GetSearchSQL()
        {
            string grade = textBoxDropDownGrade.Text;
            string _class = textBoxDropDownClass.Text;
            string name = textBoxDropDownName.Text;
            string school = textBoxDropDownSchool.Text;

            string strSql = @"select * from schoolhealth";
            bool isWhere = false;


            if (!school.IsNullOrEmpty())
            {
                strSql += " where ";
                isWhere = true;
                strSql += string.Format(" [XueXiao] like '%{0}%' ", school);
            }

            if (!grade.IsNullOrEmpty())
            {
                if (!isWhere)
                {
                    strSql += " where ";
                    isWhere = true;
                }
                else
                {
                    strSql += " and ";
                }
                strSql += string.Format(" [NianJi] like '%{0}%' ", grade);
            }

            if (!_class.IsNullOrEmpty())
            {
                if (!isWhere)
                {
                    strSql += " where ";
                    isWhere = true;
                }
                else
                {
                    strSql += " and ";
                }
                strSql += string.Format(" [BanJi] like '%{0}%' ", _class);
            }

            if (!name.IsNullOrEmpty())
            {
                if (!isWhere)
                {
                    strSql += " where ";
                    isWhere = true;
                }
                else
                {
                    strSql += " and ";
                }
                strSql += string.Format(" [XingMing] like '%{0}%' ", name);
            }
            return strSql;
        }
        #endregion

        #region Properties
        public SchoolHealthRecord Record { get; set; }
        #endregion

        #region Event Handlers

        private void dataGridViewXHistory_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridViewXHistory.SelectedRows.Count > 0)
            {
                int id = (int)dataGridViewXHistory.SelectedRows[0].Cells["colID"].Value;
                //superTabControlHistory.SelectedTabIndex = 2;
                updateControl.UpdateRecod(id);
            }
        }

        //private void btnUpdate_Click(object sender, EventArgs e)
        //{
        //    if(dataGridViewXHistory.SelectedRows.Count > 0)
        //    {
        //        int id = (int)dataGridViewXHistory.SelectedRows[0].Cells["colID"].Value;
        //        superTabControlHistory.SelectedTabIndex = 2;
        //        updateControl.UpdateRecod(id);
        //    }
        //}

        private void superTabControlMain_SelectedTabChanged(object sender, SuperTabStripSelectedTabChangedEventArgs e)
        {
            if (e.NewValue == superTabItemHistory)
            {
                RefreshSchool();
                RefreshClass();
                RefreshGrade();

                newControl.SwitchShortCut(false);
                updateControl.SwitchShortCut(true);
            }
            else if (e.NewValue == superTabItemNew)
            {
                newControl.SwitchShortCut(true);
                updateControl.SwitchShortCut(false);
            }


            if (e.NewValue == superTabItemSettings && comboBoxExPrint.Items.Count == 0)
            {
                Utility.SaveInvoke(this, () =>
                    {
                        comboBoxExPrint.SelectedIndexChanged -= comboBoxExPrint_SelectedIndexChanged;
                        comboBoxExPrint.DataSource = Utility.LoadPrinters();
                        var dp = Properties.Settings.Default.DefaultPrinter;
                        if (!dp.IsNullOrEmpty())
                        {
                            var index = comboBoxExPrint.FindStringExact(dp);
                            if (index > -1)
                            {
                                comboBoxExPrint.SelectedIndex = index;
                                comboBoxExPrint_SelectedIndexChanged(comboBoxExPrint, null);
                            }
                        }
                        else if (comboBoxExPrint.Items.Count > 0)
                        {
                            comboBoxExPrint_SelectedIndexChanged(comboBoxExPrint, null);
                        }
                        comboBoxExPrint.SelectedIndexChanged += comboBoxExPrint_SelectedIndexChanged;
                    }, true, false);

            }

            switch (superTabControlMain.SelectedTabIndex)
            {
                case 0:
                    lblCurrentSelection.Text = "新建向导";
                    break;
                case 1:
                    lblCurrentSelection.Text = "历史记录";
                    break; ;
                case 2:
                    lblCurrentSelection.Text = "系统设置";
                    break;
            }
        }

        private void textBoxDropDownSchool_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Return)
            {
                RefreshGrade();
                RefreshClass();
            }
        }

        private void textBoxDropDownSchool_ButtonClearClick(object sender, CancelEventArgs e)
        {
            RefreshGrade();
            RefreshClass();
        }

        private void textBoxDropDownGrade_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Return)
            {
                RefreshClass();
            }
        }

        private void textBoxDropDownGrade_ButtonClearClick(object sender, CancelEventArgs e)
        {
            RefreshClass();
        }

        //private void btnSave_Click(object sender, EventArgs e)
        //{
        //    if (UpdateUIInputToDataModel())
        //    {
        //        btnNew.Enabled = true;
        //        btnNewEx.Enabled = true;
        //        btnSave.Enabled = false;
        //        btnSaveEx.Enabled = false;

        //        DBHelper.Add(Record);
        //        isNewRecordAddedButNotViewHistory = true;
        //    }
        //    else
        //    {
        //        isCanProcessNewAndSave = false;
        //    }
        //}

        //private void btnNew_Click(object sender, EventArgs e)
        //{
        //    foreach (var ep in expandPanels)
        //    {
        //        ep.Enabled = true;
        //    }

        //    btnNew.Enabled = false;
        //    btnNewEx.Enabled = false;
        //    btnSave.Enabled = true;
        //    btnSaveEx.Enabled = true;
        //    btnSaveAndNew.Enabled = true;
        //    btnSaveAndNewEx.Enabled = true;


        //    if (!switchButtonExpandCollapse.Value)
        //    {
        //        switchButtonExpandCollapse.Value = true;
        //    }
        //    else
        //    {
        //        foreach (var ep in expandPanels)
        //        {
        //            ep.Expanded = true;
        //        }
        //    }


        //    //set default value
        //    ResetComboBoxSelection();
        //    InitInputValue();




        //    if (!Settings.Default.DefaultSchool.IsNullOrEmpty())
        //    {
        //        comboBoxExXueXiao.Text = Settings.Default.DefaultSchool;
        //    }
        //    if (!Settings.Default.DefaultGrade.IsNullOrEmpty())
        //    {
        //        comboBoxExNianJi.Text = Settings.Default.DefaultGrade;
        //    }

        //    if (!Settings.Default.DefaultClass.IsNullOrEmpty())
        //    {
        //        comboBoxExBanJi.Text = Settings.Default.DefaultClass;
        //    }
        //    if (Settings.Default.CheckDate.Year >= 1990)
        //    {
        //        integerInputCheckDateYear.Value = Settings.Default.CheckDate.Year;
        //        integerInputCheckDateMonth.Value = Settings.Default.CheckDate.Month;
        //        integerInputCheckDateDay.Value = Settings.Default.CheckDate.Day;
        //    }
        //    if (!Settings.Default.Checker.IsNullOrEmpty())
        //    {
        //        comboBoxExChecker.Text = Settings.Default.Checker;
        //    }


        //    comboBoxExXueXiao.Focus();
        //}

        //private void btnNewEx_Click(object sender, EventArgs e)
        //{
        //    groupPanelRecord.ScrollControlIntoView(btnNew);
        //    btnNew_Click(null, null);
        //}

        private void dataGridViewXHistory_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {

        }

        private void dataGridViewXHistory_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.ColumnIndex > 0)
            {
                DataGridViewColumn col = dataGridViewXHistory.Columns[e.ColumnIndex];



                if (col.Name == "colXingBie")
                {
                    bool flag = (bool)e.Value;
                    e.Value = flag ? "男" : "女";
                }
                else if (col.Name == "colYuHangHuJi")
                {
                    bool flag = (bool)e.Value;
                    e.Value = flag ? "是" : "否";
                }
                else if (col.Name == "colKaBa")
                {
                    bool flag = (bool)e.Value;
                    e.Value = flag ? "有" : "无";
                }

                else if (col.Name == "colFloatReserved4")
                {
                    int val;
                    int.TryParse(e.Value.ToString(), out val);
                    if ((int)val <= 10)
                    {
                        e.Value = string.Empty;
                    }
                }

                else if (col.Name == "colShenGao")
                {
                    float val;
                    float.TryParse(e.Value.ToString(), out val);

                    e.Value = val.ToString("f1");
                  
                }
            }
        }

        //private void switchButtonExpandCollapse_ValueChanged(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        groupPanelRecord.SuspendLayout();
        //        foreach (var ep in expandPanels)
        //        {
        //            ep.Expanded = switchButtonExpandCollapse.Value;
        //        }
        //    }
        //    catch
        //    {

        //    }
        //    finally
        //    {
        //        groupPanelRecord.ResumeLayout(false);
        //        groupPanelRecord.PerformLayout();
        //    }

        //}

        private void btnSearch_Click(object sender, EventArgs e)
        {

            if (dataGridViewXHistory.Columns.Count == 0)
            {
                this.bindingNavigator.AddNewRecordButton.Visible = false;
                this.bindingNavigator.DeleteButton.Visible = false;
                this.dataGridViewXHistory.AutoGenerateColumns = false;
                this.dataGridViewXHistory.Columns.AddRange(DBHelper.DefaultColumns.ToArray());
                //        isNewRecordAddedButNotViewHistory = false;       
            }


            SearchRecord(true, "正在查询...");
        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {
            if (dataGridViewXHistory.SelectedRows.Count > 0)
            {
                if (Utility.ShowQuestion("确定要删除该条记录吗?", this) == DialogResult.Yes)
                {
                    ProgressForm.QueueWorkingItem(() =>
                    {
                        DBHelper.Delete((int)dataGridViewXHistory.SelectedRows[0].Cells["colID"].Value);
                        SearchRecord(false, string.Empty);

                    }, "正在删除...", this);
                }
            }
        }

        #region 反馈单

        private void btnPrintSearchResultFanKuiDan_Click(object sender, EventArgs e)
        {
            if (dataGridViewXHistory.Rows.Count == 0)
            {

                Utility.ShowInformation("请先进行查询再选择打印反馈单.", this);

                return;
            }

            //   isPrinting = true;
            //       printResetEvent = new AutoResetEvent(false);

            curViewer = reportViewerfkd;
            curViewer.BringToFront();

            List<int> ids = new List<int>(dataGridViewXHistory.Rows.Count);
            foreach (DataGridViewRow dr in dataGridViewXHistory.Rows)
            {
                ids.Add((int)dr.Cells[0].Value);
            }

            ProgressForm pf = new ProgressForm();

            pf.Text = "正在打印,请稍后...";
            pf.SupportCancel = true;
            pf.IsReportProgress = true;


            pf.DoWork += (ss, ee) =>
                {
                    int i = 0;
                    foreach (var id in ids)
                    {
                        if (pf.CancellationPending)
                            break;
                        Utility.SaveInvoke(this, () =>
                            {
                                Utility.PrintFanKuiDan(curViewer, id);
                            }, false, false);

                        i++;
                        pf.ReportProgress((int)(i * 100f / ids.Count), i);
                        Application.DoEvents();
                    }
                };
            pf.ProgressChanged += (ss, ee) =>
                {
                    pf.Text = string.Format("正在打印 {0} / {1}", ee.UserState, ids.Count);
                };

            pf.Start(null, this);

        }


        private void btnPrintCurFanKuiDan_Click(object sender, EventArgs e)
        {
            if (dataGridViewXHistory.SelectedRows.Count > 0)
            {
                ProgressForm.QueueWorkingItem(() =>
                {

                    superTabControlHistory.SelectedTabIndex = 1;
                    curViewer = reportViewerfkd;
                    curViewer.BringToFront();

                    //    curViewer.ReportRefresh -=new CancelEventHandler(curViewer_ReportRefreshForPrint);
                    //     curViewer.ReportRefresh +=new CancelEventHandler(curViewer_ReportRefreshForPrint);

                    curViewer.RenderingComplete -= new Microsoft.Reporting.WinForms.RenderingCompleteEventHandler(reportViewer1_RenderingCompleteForPrint);
                    curViewer.RenderingComplete += new Microsoft.Reporting.WinForms.RenderingCompleteEventHandler(reportViewer1_RenderingCompleteForPrint);


                    Utility.ReportFanKuiDan(this.curViewer, (int)dataGridViewXHistory.SelectedRows[0].Cells["colID"].Value);


                }, "正在打印反馈单...", this);
            }
            else
            {
                Utility.ShowInformation("请选择一条记录再打印反馈单", this);
            }
        }


        private void btnCurRecordFanKuiDan_Click(object sender, EventArgs e)
        {
            if (dataGridViewXHistory.SelectedRows.Count > 0)
            {
                ProgressForm.QueueWorkingItem(() =>
                    {

                        superTabControlHistory.SelectedTabIndex = 1;
                        reportViewerfkd.BringToFront();
                        Utility.ReportFanKuiDan(this.reportViewerfkd, (int)dataGridViewXHistory.SelectedRows[0].Cells["colID"].Value);


                    }, "正在生成反馈单...", this);
            }
            else
            {
                Utility.ShowInformation("请选择一条记录再生成反馈单", this);
            }
        }

        private void btnExportCurFanKuiDan_Click(object sender, EventArgs e)
        {
            if (dataGridViewXHistory.SelectedRows.Count > 0)
            {
                ProgressForm.QueueWorkingItem(() =>
                {

                    superTabControlHistory.SelectedTabIndex = 1;
                    this.expType = ExportType.Word;

                    curViewer = reportViewerfkd;
                    curViewer.BringToFront();
                    //     curViewer.ReportRefresh -= new CancelEventHandler(curViewer_ReportRefreshForExport);
                    //      curViewer.ReportRefresh += new CancelEventHandler(curViewer_ReportRefreshForExport);

                    curViewer.RenderingComplete -= new Microsoft.Reporting.WinForms.RenderingCompleteEventHandler(reportViewer1_RenderingCompleteForExport);
                    curViewer.RenderingComplete += new Microsoft.Reporting.WinForms.RenderingCompleteEventHandler(reportViewer1_RenderingCompleteForExport);

                    Utility.ReportFanKuiDan(this.curViewer, (int)dataGridViewXHistory.SelectedRows[0].Cells["colID"].Value);



                }, "正在导出反馈单...", this);
            }
            else
            {
                Utility.ShowInformation("请选择一条记录再导出反馈单", this);
            }
        }
        #endregion

        #region 体检结果
        private void btnTiJianJieGuoBySearch_Click(object sender, EventArgs e)
        {
            string s = textBoxDropDownSchool.Text;
            ProgressForm.QueueWorkingItem(() =>
            {
                superTabControlHistory.SelectedTabIndex = 1;
                string strSql = GetSearchSQL();

                var dataSet = DBHelper.SchoolHealthQueryDataSet(strSql);
                var dt = DBHelper.TiJianJieGuoTable(dataSet.Tables[0], s);
                reportViewertjjg.BringToFront();

                Utility.ReportTiJianJieGuo(reportViewertjjg, dt);

            }, "正在生成体检结果...", this);
        }

        private void btnExportTiJianJieGuoBySearch_Click(object sender, EventArgs e)
        {

            string s = textBoxDropDownSchool.Text;
            ProgressForm.QueueWorkingItem(() =>
            {
                superTabControlHistory.SelectedTabIndex = 1;
                string strSql = GetSearchSQL();
                this.expType = ExportType.Excel;
                var dataSet = DBHelper.SchoolHealthQueryDataSet(strSql);
                var dt = DBHelper.TiJianJieGuoTable(dataSet.Tables[0], s);
                curViewer = reportViewertjjg;
                curViewer.BringToFront();
                //   curViewer.ReportRefresh -= new CancelEventHandler(curViewer_ReportRefreshForExport);
                //   curViewer.ReportRefresh += new CancelEventHandler(curViewer_ReportRefreshForExport);


                curViewer.RenderingComplete -= new Microsoft.Reporting.WinForms.RenderingCompleteEventHandler(reportViewer1_RenderingCompleteForExport);
                curViewer.RenderingComplete += new Microsoft.Reporting.WinForms.RenderingCompleteEventHandler(reportViewer1_RenderingCompleteForExport);
                Utility.ReportTiJianJieGuo(curViewer, dt);



            }, "正在导出体检结果...", this);
        }

        private void btnTiJianJieGuoAll_Click(object sender, EventArgs e)
        {
            if (textBoxDropDownSchool.Text.IsNullOrEmpty() && textBoxDropDownSchool.DropDownItems.Count > 1)
            {
                Utility.ShowInformation("存在多个学校的记录,请先按学校进行查询,再生成体检结果.", this);
                textBoxDropDownSchool.Focus();
                return;
            }

            string s = textBoxDropDownSchool.Text;

            ProgressForm.QueueWorkingItem(() =>
            {
                superTabControlHistory.SelectedTabIndex = 1;
                var dt = DBHelper.TiJianJieGuoTable(null, s);
                reportViewertjjg.BringToFront();
                Utility.ReportTiJianJieGuo(reportViewertjjg, dt, textBoxDropDownSchool.Text);


            }, "正在生成体检结果...", this);
        }

        private void btnExportTiJianJieGuoAll_Click(object sender, EventArgs e)
        {
            if (textBoxDropDownSchool.Text.IsNullOrEmpty() && textBoxDropDownSchool.DropDownItems.Count > 1)
            {
                Utility.ShowInformation("存在多个学校的记录,请先按学校进行查询,再导出体检结果.", this);
                textBoxDropDownSchool.Focus();
                return;
            }

            string s = textBoxDropDownSchool.Text;

            ProgressForm.QueueWorkingItem(() =>
            {
                superTabControlHistory.SelectedTabIndex = 1;
                var dt = DBHelper.TiJianJieGuoTable(null, s);
                this.expType = ExportType.Excel;
                curViewer = reportViewertjjg;
                curViewer.BringToFront();
                //  curViewer.ReportRefresh -= new CancelEventHandler(curViewer_ReportRefreshForPrint);
                //   curViewer.ReportRefresh += new CancelEventHandler(curViewer_ReportRefreshForPrint);

                curViewer.RenderingComplete -= new Microsoft.Reporting.WinForms.RenderingCompleteEventHandler(reportViewer1_RenderingCompleteForExport);
                curViewer.RenderingComplete += new Microsoft.Reporting.WinForms.RenderingCompleteEventHandler(reportViewer1_RenderingCompleteForExport);


                Utility.ReportTiJianJieGuo(curViewer, dt);

            }, "正在导出体检结果...", this);
        }
        #endregion

        #region 体检汇总
        private void btnExportTiJianHuiZong_Click(object sender, EventArgs e)
        {
            //if (textBoxDropDownSchool.Text.IsNullOrEmpty() && textBoxDropDownSchool.DropDownItems.Count > 1)
            //{
            //    Utility.ShowInformation("存在多个学校的记录,请先按学校进行查询,再进行汇总.", this);
            //    textBoxDropDownSchool.Focus();
            //    return;
            //}

            var s = textBoxDropDownSchool.Text;

            ProgressForm.QueueWorkingItem(() =>
            {
                superTabControlHistory.SelectedTabIndex = 1;
                this.expType = ExportType.Excel;

                curViewer = reportViewertjfz;
                curViewer.BringToFront();
                //  curViewer.ReportRefresh -= new CancelEventHandler(curViewer_ReportRefreshForExport);
                //     curViewer.ReportRefresh += new CancelEventHandler(curViewer_ReportRefreshForExport);

                curViewer.RenderingComplete -= new Microsoft.Reporting.WinForms.RenderingCompleteEventHandler(reportViewer1_RenderingCompleteForExport);
                curViewer.RenderingComplete += new Microsoft.Reporting.WinForms.RenderingCompleteEventHandler(reportViewer1_RenderingCompleteForExport);

                Utility.ReportTiJianHuiZong(curViewer, s);

            }, "正在导出体检汇总...", this);

        }

        private void btnTiJianHuiZong_Click(object sender, EventArgs e)
        {
            //if (textBoxDropDownSchool.Text.IsNullOrEmpty() && textBoxDropDownSchool.DropDownItems.Count > 1)
            //{
            //    Utility.ShowInformation("存在多个学校的记录,请先按学校进行查询,再进行汇总.", this);
            //    textBoxDropDownSchool.Focus();
            //    return;
            //}

            var s = textBoxDropDownSchool.Text;

            ProgressForm.QueueWorkingItem(() =>
            {
                superTabControlHistory.SelectedTabIndex = 1;
                reportViewertjfz.BringToFront();
                Utility.ReportTiJianHuiZong(reportViewertjfz, s);

            }, "正在生成体检汇总...", this);
        }
        #endregion


        //void curViewer_ReportRefreshForExport(object sender, CancelEventArgs e)
        //{
        //    Utility.ExportReport(curViewer, expType);
        //    curViewer.ReportRefresh -= curViewer_ReportRefreshForExport;
        //}


        //void curViewer_ReportRefreshForPrint(object sender, CancelEventArgs e)
        //{
        //    Utility.PrintReport(this.curViewer, this);
        //    curViewer.ReportRefresh -= curViewer_ReportRefreshForPrint;
        //}

        void reportViewer1_RenderingCompleteForExport(object sender, Microsoft.Reporting.WinForms.RenderingCompleteEventArgs e)
        {
            Utility.ExportReport(curViewer, expType);
            curViewer.RenderingComplete -= new Microsoft.Reporting.WinForms.RenderingCompleteEventHandler(reportViewer1_RenderingCompleteForExport);
        }

        void reportViewer1_RenderingCompleteForPrint(object sender, Microsoft.Reporting.WinForms.RenderingCompleteEventArgs e)
        {
            if (this.InvokeRequired)
            {
                Utility.SaveInvoke(this, () =>
                    {
                        reportViewer1_RenderingCompleteForPrint(sender, e);
                    }, false, false);
                return;
            }
            Utility.PrintReport(this.curViewer, this, false);

            //if(isPrinting && printOneReportEvent != null)
            //{
            //    printOneReportEvent.Set();
            //}
        }

        private void superTabControlHistory_SelectedTabChanged(object sender, SuperTabStripSelectedTabChangedEventArgs e)
        {
            btnDelete.Visible = superTabControlHistory.SelectedTabIndex == 0;
        }
        private void comboBoxExPrint_SelectedIndexChanged(object sender, EventArgs e)
        {
            Utility.DefaultPrinter = comboBoxExPrint.SelectedItem.ToString();
        }
        private void comboBoxExUIStyle_SelectedIndexChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.DefaultStyle = comboBoxExUIStyle.SelectedItem.ToString();
            styleManager1.ManagerStyle = (eStyle)Enum.Parse(typeof(eStyle), Properties.Settings.Default.DefaultStyle);
            Properties.Settings.Default.Save();
        }

        #endregion

        private void btnBackup_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Filter = "数据备份(*.mdb)|*.mdb";

                dlg.FileName = string.Format("{0:yyyy年MM月dd日HH时mm分ss秒}备份.mdb", DateTime.Now);
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    if (DBHelper.Backup(dlg.FileName))
                    {
                        Dictionary<string, string> dic = new Dictionary<string, string>();
                        dic[dlg.FileName] = dlg.FileName;

                        Utility.ShowInformation("数据备份成功");
                        Utility.OpenFileLocations(dic);
                    }
                    else
                    {
                        Utility.ShowInformation("未知原因,数据备份失败");
                    }
                }
            }
        }

        private void btnRestore_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "数据备份(*.mdb)|*.mdb";

                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    DBHelper.Restore(dlg.FileName);
                }
            }
        }

        private void checkBoxMMHG_CheckedChanged(object sender, EventArgs e)
        {
            Settings.Default.IsMMHG = checkBoxMMHG.Checked;
            Settings.Default.Save();

        }

        private void btnUpdateStatisticResult_Click(object sender, EventArgs e)
        {
            DBHelper.UpdateStatisticResult();
        }







        //private void btnSaveAndNew_Click(object sender, EventArgs e)
        //{
        //    isCanProcessNewAndSave = true;
        //    btnSave_Click(null, null);
        //    if (!isCanProcessNewAndSave) return;
        //    btnNewEx_Click(null, null);
        //}


    }
}