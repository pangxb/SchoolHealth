using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using DevComponents.DotNetBar.Controls;
using System.Windows.Forms;
using DevComponents.DotNetBar;

namespace SchoolHealth
{
    public static class ConfigurationHelper
    {    
        public static Dictionary<string,List<string>> AppSettings {private set;get;}

        public static readonly string 性别 = "性别";
        public static readonly string 余杭户籍 = "余杭户籍";
        public static readonly string 心 = "心";
        public static readonly string 年级 = "年级";
        public static readonly string 肺 = "肺";
        public static readonly string 肝 = "肝";
        public static readonly string 脾 = "脾";
        public static readonly string 沙眼 = "沙眼";
        public static readonly string 结膜炎 = "结膜炎";
        public static readonly string 色觉 = "色觉";
        public static readonly string 牙周 = "牙周";
        public static readonly string 耳 = "耳";
        public static readonly string 鼻 = "鼻";
        public static readonly string 扁桃体 = "扁桃体";
        public static readonly string 头部 = "头部";
        public static readonly string 颈部 = "颈部";
        public static readonly string 胸部 = "胸部";
        public static readonly string 脊柱 = "脊柱";
        public static readonly string 四肢 = "四肢";
        public static readonly string 皮肤 = "皮肤";
        public static readonly string 淋巴结 = "淋巴结";

        public static readonly string 血型 = "血型";
        public static readonly string 蛔虫卵 = "肠道蠕虫卵";
        public static readonly string 卡疤 = "卡疤";
        public static readonly string 结核菌素试验 = "结核菌素试验";
        public static readonly string 谷丙转氨酶 = "谷丙转氨酶";
        public static readonly string 胆红素 = "胆红素";
        public static readonly string Excel数据类型 = "Excel数据类型";
        

        static ConfigurationHelper()
        {
            AppSettings = new Dictionary<string, List<string>>();          

            foreach (var key in ConfigurationManager.AppSettings.AllKeys)
            {
                string val = ConfigurationManager.AppSettings[key];
                string[] values = val.Split(',');
                AppSettings[key] = new List<string>(values);
            }
        }

        public static List<String> DataSourceByKey(string key)
        {
            List<string> source;
            if (AppSettings.TryGetValue(key, out source))
            {
                return source;
            }
            else 
            {
                return new List<string>();
            }
        }

        public static void BindDataSource(string key, ComboBoxEx cbx)
        {
            cbx.DataSource = DataSourceByKey(key);
            cbx.KeyUp +=(s,e)=>
            {
                int k = (int)e.KeyCode;
              
                int key0 = (int)Keys.D0;
                int key9 = (int)Keys.D9;

                if (k >= key0 && k <= key9 && cbx.Items.Count >= (k - key0))
                {
                    cbx.SelectedIndex = (k - key0-1);
                    return;
                }

             

                 key0 = (int)Keys.NumPad0;
                 key9 = (int)Keys.NumPad9;

                 if (k >= key0 && k <= key9 && cbx.Items.Count >= (k - key0))
                {
                    cbx.SelectedIndex = (k - key0-1);
                }


            };
        }

        public static void BindDataSource(string key, TextBoxDropDown cbx, MethodInvoker callback = null)
        {
     
            cbx.AutoCompleteCustomSource.AddRange(DataSourceByKey(key).ToArray());
            foreach (var s in DataSourceByKey(key))
            {
                ButtonItem bi = new ButtonItem("",s);
                bi.Click += delegate
                {
                    cbx.Text = bi.Text;
                    if(callback != null)
                    {
                        callback();
                    }
                };
                cbx.DropDownItems.Add(bi);
            }
        }
    }
}
