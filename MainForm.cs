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
                        Utility.ShowInformation("��Ҫ����");
                        Settings.Default.UpdateFlag = 2;
                        Settings.Default.Save();
#else              
                        DBHelper.UpdateStatisticResult(true);             
#endif

                    }
                    else
                    {
#if DEBUG
                        Utility.ShowInformation("����Ҫ����");
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
            // startUpHelper.UpdateMsg("���ڳ�ʼ������Դ");
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

                //��ӡ��
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

                var defaultGrades = ConfigurationHelper.DataSourceByKey(ConfigurationHelper.�꼶);

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
        //        }, "����ˢ������...", this);


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
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.�꼶, comboBoxExNianJi);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.�Ա�, comboBoxExXingBie);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.�ຼ����, comboBoxExYuHang);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.��, comboBoxExXin);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.��, comboBoxExFei);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.��, comboBoxExGan);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.Ƣ, comboBoxExPi);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.ɳ��, comboBoxExShaYan);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.��Ĥ��, comboBoxExJieMoYan);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.ɫ��, comboBoxExSeJue);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.����, comboBoxExYaZhou);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.��, comboBoxExEr);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.��, comboBoxExBi);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.������, comboBoxExBianTaoTi);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.ͷ��, comboBoxExTou);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.����, comboBoxExJing);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.�ز�, comboBoxExXiong);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.����, comboBoxExJiZhu);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.��֫, comboBoxExSiZhi);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.Ƥ��, comboBoxExPiFu);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.�ܰͽ�, comboBoxExLinBaJie);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.Ѫ��, comboBoxExXueXing);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.�׳���, comboBoxExHuiChongLuan);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.����, comboBoxExKaBa);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.��˾�������, comboBoxExJieHeJun);

        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.�ȱ�ת��ø, comboBoxExZhuanAnMei);
        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.������, comboBoxExDanHongSu);

        //    ConfigurationHelper.BindDataSource(ConfigurationHelper.�꼶, textBoxDropDownGrade, RefreshClass);
        //    // ConfigurationHelper.BindDataSource(ConfigurationHelper.Excel��������, comboBoxExExcelExportType);


        //    //��ӡ��
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
        //        Utility.ShowInformation("ѧУ����Ϊ��!", this);
        //        ScrollToVisible(comboBoxExXueXiao);
        //        return false;
        //    }

        //    if (comboBoxExBanJi.Text.IsNullOrEmpty())
        //    {
        //        Utility.ShowInformation("�༶����Ϊ��!", this);
        //        ScrollToVisible(comboBoxExBanJi);
        //        return false;
        //    }

        //    if (textBoxXXingMing.Text.IsNullOrEmpty())
        //    {
        //        Utility.ShowInformation("��������Ϊ��!", this);
        //        ScrollToVisible(textBoxXXingMing);
        //        return false;
        //    }

        //    if (comboBoxExChecker.Text.IsNullOrEmpty())
        //    {
        //        Utility.ShowInformation("��쵥λ����Ϊ��!", this);
        //        ScrollToVisible(comboBoxExChecker);
        //        return false;
        //    }

        //    if ((int)doubleInputHeight.Value == 0 ||  (int)doubleInputHeight.Value == (int)doubleInputHeight.MinValue)
        //    {
        //        Utility.ShowInformation("���������!", this);
        //        ScrollToVisible(doubleInputHeight);
        //        return false;
        //    }

        //    if ((int)doubleInputWeight.Value == 0 || (int)doubleInputWeight.Value == (int)doubleInputWeight.MinValue)
        //    {
        //        Utility.ShowInformation("����������!", this);
        //        ScrollToVisible(doubleInputWeight);
        //        return false;
        //    }

        //    //if (integerInputLungCapacity.Value == 0)
        //    //{
        //    //    Utility.ShowInformation("������λ���!", this);
        //    //    ScrollToVisible(integerInputLungCapacity);
        //    //    return false;
        //    //}

        //    if (integerInputShuZhangYa.Value == 0 || integerInputShuZhangYa.Value == integerInputShuZhangYa.MinValue)
        //    {
        //        Utility.ShowInformation("����������ѹ!", this);
        //        ScrollToVisible(integerInputShuZhangYa);
        //        return false;
        //    }

        //    if (integerInputSouSuoYa.Value == 0 || integerInputSouSuoYa.Value == integerInputSouSuoYa.MinValue)
        //    {
        //        Utility.ShowInformation("����������ѹ!", this);
        //        ScrollToVisible(integerInputSouSuoYa);
        //        return false;
        //    }

        //    if (integerInputMaiBo.Value == 0 || integerInputMaiBo.Value == integerInputMaiBo.MinValue)
        //    {
        //        Utility.ShowInformation("����������!", this);
        //        ScrollToVisible(integerInputMaiBo);
        //        return false;
        //    }


        //    //if ((int)doubleInputXueHongDanBai.Value == 0)
        //    //{
        //    //    Utility.ShowInformation("������Ѫ�쵰��ֵ!", this);
        //    //    ScrollToVisible(doubleInputXueHongDanBai);
        //    //    return false;
        //    //}


        //    //assign 
        //    //1) ������Ϣ
        //    Record.XueXiao = comboBoxExXueXiao.Text;
        //    Record.BanJi = comboBoxExBanJi.Text;
        //    Record.NianJi = comboBoxExNianJi.Text;
        //    Record.XingMing = textBoxXXingMing.Text;
        //    Record.XingBie = comboBoxExXingBie.SelectedIndex == 0;
        //    Record.ShengRi = new DateTime(integerInputBirthDateYear.Value, integerInputBirthDateMonth.Value, integerInputBirthDateDay.Value);
        //    Record.YuHangHuJi = comboBoxExYuHang.SelectedIndex == 0;
        //    Record.TiJianRiQi = new DateTime(integerInputCheckDateYear.Value, integerInputCheckDateMonth.Value, integerInputCheckDateDay.Value);
        //    Record.TiJianDanWei = comboBoxExChecker.Text;

        //    //2) ������ʷ
        //    Record.GanYan = checkBoxXGanSick.Checked;
        //    Record.FeiJieHe = checkBoxXFeiSick.Checked;
        //    Record.XinZangBing = checkBoxXXinSick.Checked;
        //    Record.ShenYan = checkBoxXShenSick.Checked;
        //    Record.FengShiBing = checkBoxXFengShiSick.Checked;
        //    Record.DiFangBing = comboBoxExLocalSick.Text;
        //    Record.QiTaBing = comboBoxExOtherSick.Text;
        //    Record.ShengBingRiQi = new DateTime(integerInputSickDateYear.Value, integerInputSickDateMonth.Value, integerInputSickDateDay.Value);

        //    //3)�������
        //    Record.ShenGao = doubleInputHeight.Value.ToFloat();
        //    Record.TiZhong = doubleInputWeight.Value.ToFloat();
        //    Record.FeiHuoLiang = integerInputLungCapacity.Value;
        //    Record.SuZhangYa = integerInputShuZhangYa.Value;
        //    Record.SouSuoYa = integerInputSouSuoYa.Value;
        //    Record.MaiBo = integerInputMaiBo.Value;

        //    //4)�ڿ�
        //    Record.Xin = comboBoxExXin.Text;
        //    Record.Fei = comboBoxExFei.Text;
        //    Record.Gan = comboBoxExGan.Text;
        //    Record.Pi = comboBoxExPi.Text;

        //    //5)�ۿ�
        //    Record.ZuoYan = doubleInputZuoYan.Value.ToFloat();
        //    Record.YouYan = doubleInputYouYan.Value.ToFloat();
        //    Record.ShaYan = comboBoxExShaYan.Text;
        //    Record.JieMoYan = comboBoxExJieMoYan.Text;
        //    Record.SeJue = comboBoxExSeJue.Text;

        //    //6)��ǻ��
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

        //    //7)�����ʺ��
        //    Record.Er = comboBoxExEr.Text;
        //    Record.Bi = comboBoxExBi.Text;
        //    Record.BianTaoTi = comboBoxExBianTaoTi.Text;

        //    //8)���
        //    Record.Tou = comboBoxExTou.Text;
        //    Record.Jing = comboBoxExJing.Text;
        //    Record.Xiong = comboBoxExXiong.Text;
        //    Record.JiZhu = comboBoxExJiZhu.Text;
        //    Record.SiZhi = comboBoxExSiZhi.Text;
        //    Record.PiFu = comboBoxExPiFu.Text;
        //    Record.LinBaJie = comboBoxExLinBaJie.Text;

        //    //9)Ѫ��
        //    Record.XueXing = comboBoxExXueXing.Text;
        //    Record.XueHongDanBai = doubleInputXueHongDanBai.Value.ToFloat();
        //    Record.HuiChongLuan = comboBoxExHuiChongLuan.Text;

        //    //10)�ν��
        //    Record.KaBa = comboBoxExKaBa.SelectedIndex == 0;
        //    Record.JieHeJunSu = comboBoxExJieHeJun.Text;

        //    //11)�ι���
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
        //    //1) ������Ϣ
        //    comboBoxExXueXiao.Text = Record.XueXiao;
        //    comboBoxExBanJi.Text = Record.BanJi;
        //    comboBoxExNianJi.Text = Record.NianJi;
        //    //Record.XingMing = textBoxXXingMing.Text;
        //    //Record.XingBie comboBoxExXingBie.SelectedIndex == 0;
        //    //Record.ShengRi = new DateTime(integerInputBirthDateYear.Value, integerInputBirthDateMonth.Value, integerInputBirthDateDay.Value);
        //    //Record.YuHangHuJi comboBoxExYuHang.SelectedIndex == 0;
        //    //Record.TiJianRiQi = new DateTime(integerInputCheckDateYear.Value, integerInputCheckDateMonth.Value, integerInputCheckDateDay.Value);
        //    //Record.TiJianDanWei comboBoxExChecker.Text;

        //    ////2) ������ʷ
        //    //Record.GanYan = checkBoxXGanSick.Checked;
        //    //Record.FeiJieHe = checkBoxXFeiSick.Checked;
        //    //Record.XinZangBing = checkBoxXXinSick.Checked;
        //    //Record.ShenYan = checkBoxXShenSick.Checked;
        //    //Record.FengShiBing = checkBoxXFengShiSick.Checked;
        //    //Record.DiFangBing comboBoxExLocalSick.Text;
        //    //Record.QiTaBing comboBoxExOtherSick.Text;
        //    //Record.ShengBingRiQi = new DateTime(integerInputSickDateYear.Value, integerInputSickDateMonth.Value, integerInputSickDateDay.Value);

        //    ////3)�������
        //    //Record.ShenGao = doubleInputHeight.Value.ToFloat();
        //    //Record.TiZhong = doubleInputWeight.Value.ToFloat();
        //    //Record.FeiHuoLiang = integerInputLungCapacity.Value;
        //    //Record.SuZhangYa = integerInputShuZhangYa.Value;
        //    //Record.SouSuoYa = integerInputSouSuoYa.Value;
        //    //Record.MaiBo = integerInputMaiBo.Value;

        //    ////4)�ڿ�
        //    //Record.Xin comboBoxExXin.Text;
        //    //Record.Fei comboBoxExFei.Text;
        //    //Record.Gan comboBoxExGan.Text;
        //    //Record.Pi comboBoxExPi.Text;

        //    ////5)�ۿ�
        //    //Record.ZuoYan = doubleInputZuoYan.Value.ToFloat();
        //    //Record.YouYan = doubleInputYouYan.Value.ToFloat();
        //    //Record.ShaYan comboBoxExShaYan.Text;
        //    //Record.JieMoYan comboBoxExJieMoYan.Text;
        //    //Record.SeJue comboBoxExSeJue.Text;

        //    ////6)��ǻ��
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

        //    ////7)�����ʺ��
        //    //Record.Er comboBoxExEr.Text;
        //    //Record.Bi comboBoxExBi.Text;
        //    //Record.BianTaoTi comboBoxExBianTaoTi.Text;

        //    ////8)���
        //    //Record.Tou comboBoxExTou.Text;
        //    //Record.Jing comboBoxExJing.Text;
        //    //Record.Xiong comboBoxExXiong.Text;
        //    //Record.JiZhu comboBoxExJiZhu.Text;
        //    //Record.SiZhi comboBoxExSiZhi.Text;
        //    //Record.PiFu comboBoxExPiFu.Text;
        //    //Record.LinBaJie comboBoxExLinBaJie.Text;

        //    ////9)Ѫ��
        //    //Record.XueXing comboBoxExXueXing.Text;
        //    //Record.XueHongDanBai = doubleInputXueHongDanBai.Value.ToFloat();
        //    //Record.HuiChongLuan comboBoxExHuiChongLuan.Text;

        //    ////10)�ν��
        //    //Record.KaBa comboBoxExKaBa.SelectedIndex == 0;
        //    //Record.JieHeJunShiYan comboBoxExJieHeJun.Text;

        //    ////11)�ι���
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
                    lblCurrentSelection.Text = "�½���";
                    break;
                case 1:
                    lblCurrentSelection.Text = "��ʷ��¼";
                    break; ;
                case 2:
                    lblCurrentSelection.Text = "ϵͳ����";
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
                    e.Value = flag ? "��" : "Ů";
                }
                else if (col.Name == "colYuHangHuJi")
                {
                    bool flag = (bool)e.Value;
                    e.Value = flag ? "��" : "��";
                }
                else if (col.Name == "colKaBa")
                {
                    bool flag = (bool)e.Value;
                    e.Value = flag ? "��" : "��";
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


            SearchRecord(true, "���ڲ�ѯ...");
        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {
            if (dataGridViewXHistory.SelectedRows.Count > 0)
            {
                if (Utility.ShowQuestion("ȷ��Ҫɾ��������¼��?", this) == DialogResult.Yes)
                {
                    ProgressForm.QueueWorkingItem(() =>
                    {
                        DBHelper.Delete((int)dataGridViewXHistory.SelectedRows[0].Cells["colID"].Value);
                        SearchRecord(false, string.Empty);

                    }, "����ɾ��...", this);
                }
            }
        }

        #region ������

        private void btnPrintSearchResultFanKuiDan_Click(object sender, EventArgs e)
        {
            if (dataGridViewXHistory.Rows.Count == 0)
            {

                Utility.ShowInformation("���Ƚ��в�ѯ��ѡ���ӡ������.", this);

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

            pf.Text = "���ڴ�ӡ,���Ժ�...";
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
                    pf.Text = string.Format("���ڴ�ӡ {0} / {1}", ee.UserState, ids.Count);
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


                }, "���ڴ�ӡ������...", this);
            }
            else
            {
                Utility.ShowInformation("��ѡ��һ����¼�ٴ�ӡ������", this);
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


                    }, "�������ɷ�����...", this);
            }
            else
            {
                Utility.ShowInformation("��ѡ��һ����¼�����ɷ�����", this);
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



                }, "���ڵ���������...", this);
            }
            else
            {
                Utility.ShowInformation("��ѡ��һ����¼�ٵ���������", this);
            }
        }
        #endregion

        #region �����
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

            }, "�������������...", this);
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



            }, "���ڵ��������...", this);
        }

        private void btnTiJianJieGuoAll_Click(object sender, EventArgs e)
        {
            if (textBoxDropDownSchool.Text.IsNullOrEmpty() && textBoxDropDownSchool.DropDownItems.Count > 1)
            {
                Utility.ShowInformation("���ڶ��ѧУ�ļ�¼,���Ȱ�ѧУ���в�ѯ,�����������.", this);
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


            }, "�������������...", this);
        }

        private void btnExportTiJianJieGuoAll_Click(object sender, EventArgs e)
        {
            if (textBoxDropDownSchool.Text.IsNullOrEmpty() && textBoxDropDownSchool.DropDownItems.Count > 1)
            {
                Utility.ShowInformation("���ڶ��ѧУ�ļ�¼,���Ȱ�ѧУ���в�ѯ,�ٵ��������.", this);
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

            }, "���ڵ��������...", this);
        }
        #endregion

        #region ������
        private void btnExportTiJianHuiZong_Click(object sender, EventArgs e)
        {
            //if (textBoxDropDownSchool.Text.IsNullOrEmpty() && textBoxDropDownSchool.DropDownItems.Count > 1)
            //{
            //    Utility.ShowInformation("���ڶ��ѧУ�ļ�¼,���Ȱ�ѧУ���в�ѯ,�ٽ��л���.", this);
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

            }, "���ڵ���������...", this);

        }

        private void btnTiJianHuiZong_Click(object sender, EventArgs e)
        {
            //if (textBoxDropDownSchool.Text.IsNullOrEmpty() && textBoxDropDownSchool.DropDownItems.Count > 1)
            //{
            //    Utility.ShowInformation("���ڶ��ѧУ�ļ�¼,���Ȱ�ѧУ���в�ѯ,�ٽ��л���.", this);
            //    textBoxDropDownSchool.Focus();
            //    return;
            //}

            var s = textBoxDropDownSchool.Text;

            ProgressForm.QueueWorkingItem(() =>
            {
                superTabControlHistory.SelectedTabIndex = 1;
                reportViewertjfz.BringToFront();
                Utility.ReportTiJianHuiZong(reportViewertjfz, s);

            }, "��������������...", this);
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
                dlg.Filter = "���ݱ���(*.mdb)|*.mdb";

                dlg.FileName = string.Format("{0:yyyy��MM��dd��HHʱmm��ss��}����.mdb", DateTime.Now);
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    if (DBHelper.Backup(dlg.FileName))
                    {
                        Dictionary<string, string> dic = new Dictionary<string, string>();
                        dic[dlg.FileName] = dlg.FileName;

                        Utility.ShowInformation("���ݱ��ݳɹ�");
                        Utility.OpenFileLocations(dic);
                    }
                    else
                    {
                        Utility.ShowInformation("δ֪ԭ��,���ݱ���ʧ��");
                    }
                }
            }
        }

        private void btnRestore_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "���ݱ���(*.mdb)|*.mdb";

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