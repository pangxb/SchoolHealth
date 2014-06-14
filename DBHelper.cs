using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Data;
using System.Linq;


using System.Data.Common;
using System.Windows.Forms;
using SchoolHealth.Properties;
using Microsoft.Reporting.WinForms;
using System.Reflection;
using System.IO;



namespace SchoolHealth
{
    static class DBHelper
    {

        // Fields
        public static OleDbConnection conn = new OleDbConnection(Settings.Default.DataConnectionString);

        private static void OpenConnectionIfNecessary()
        {
            //open connection
            if(conn.State == ConnectionState.Closed)
            {
                conn.Open();
            }
        }
        // Methods
        public static OleDbDataReader SchoolHealthDataReader(string strSql)
        {
            OleDbDataReader OleDbdr = null;
            string cmdText = strSql;
            OleDbCommand cmd = new OleDbCommand(cmdText, conn);
            OpenConnectionIfNecessary();
            try
            {
                OleDbdr = cmd.ExecuteReader();
            }
            catch (Exception ex)
            {
                Utility.ShowInformation("操作发生错误:\r\n" + ex.Message);
            }
            return OleDbdr;
        }

        public static bool SchoolHealthNonQueryCommand(string strSql)
        {
            if (strSql.Trim() != string.Empty)
            {
                OleDbCommand cmd = new OleDbCommand(strSql.Trim(), conn);
                OpenConnectionIfNecessary();
                try
                {
                    return cmd.ExecuteNonQuery() > 0;
                }
                catch (Exception ex)
                {                       
                       Utility.ShowInformation("操作发生错误:\r\n"+ex.Message);                        
                }
                finally
                {
                    cmd.Connection.Close();
                }
            }
            return false;
        }

        public static DataSet SchoolHealthQueryDataSet(string strSql)
        {
            DataSet ds = new DataSet();
            if (strSql.Trim() != string.Empty)
            {
                try
                {
                    OpenConnectionIfNecessary();
                    new OleDbDataAdapter(strSql, conn).Fill(ds, "SchoolHealth");
                }
                   catch (Exception ex)
                {              
                        ds = null;
                       Utility.ShowInformation("操作发生错误:\r\n"+ex.Message);                        
                }
            }
            return ds;
        }

        public static bool SchoolHealthUpdateDataSet(string strSql, DataSet changdataset)
        {
            DataSet ds = new DataSet();
            if (strSql.Trim() != string.Empty)
            {
                try
                {
                    OpenConnectionIfNecessary();
                    new OleDbDataAdapter(strSql, conn).Update(changdataset);
                    changdataset.AcceptChanges();
                    return true;
                }
                catch (Exception ex)
                {
                    Utility.ShowInformation("操作发生错误:\r\n" + ex.Message);
                    return false;
                }
            }
            return false;
        }

        public static List<string> ListSchool()
        {
            List<string> ret = new List<string>();
            OleDbDataReader reader = null;

            reader = SchoolHealthDataReader(@"select distinct [XueXiao] from schoolhealth");


            if (reader != null)
            {
                while (reader.Read())
                {
                    ret.Add(reader[0].ToString());
                }
                reader.Close();
                reader.Dispose();
            }
            return ret;
        }


        public static List<string> ListGrade(string school = null)
        {
            List<string> ret = new List<string>();
            OleDbDataReader reader = null;

            if (school.IsNullOrEmpty())
            {
                reader = SchoolHealthDataReader(@"select distinct [NianJi] from schoolhealth");
            }
            else
            {
                reader = SchoolHealthDataReader(string.Format(@"select distinct [NianJi] from schoolhealth where [XueXiao] = '{0}'", school));
            }
          

            if (reader != null)
            {
                while (reader.Read())
                {
                    ret.Add(reader[0].ToString());
                }
                reader.Close();
                reader.Dispose();
            }
            return ret;
        }

        public static List<string> ListClass(string school = null, string grade = null)
        {
            List<string> ret = new List<string>();
            OleDbDataReader reader = null;

            if(school.IsNullOrEmpty() && grade.IsNullOrEmpty())
            {
                reader = SchoolHealthDataReader(string.Format("select distinct [BanJi] from schoolhealth"));
            }
            else if(!school.IsNullOrEmpty() && grade.IsNullOrEmpty())
            {
                reader = SchoolHealthDataReader(string.Format("select distinct [BanJi] from schoolhealth where [XueXiao] = '{0}'", school));
            }
            else if(school.IsNullOrEmpty() && !grade.IsNullOrEmpty())
            {
                reader = SchoolHealthDataReader(string.Format("select distinct [BanJi] from schoolhealth where [NianJi] = '{0}'", grade));
            }
            else 
            {
                reader = SchoolHealthDataReader(string.Format("select distinct [BanJi] from schoolhealth where [NianJi] = '{0}' and [XueXiao] = '{1}'", grade,school));

            }
         

            if(reader != null)
            {
                while(reader.Read())
                {
                    ret.Add(reader[0].ToString());
                }
                reader.Close();
                reader.Dispose();
            }
            return ret;
        }


        public static void Add(SchoolHealthRecord record, bool assignGuid =  true)
        {
            if (assignGuid)
            {
                record.GUID = Guid.NewGuid().ToString();
                UpdateStaticResult(record);
            }
            string strSql = CreateInsertCommandText(record);
            SchoolHealthNonQueryCommand(strSql);
        }

        private static void UpdateStaticResult(SchoolHealthRecord record)
        {
            //        public String ShiLiJieGuo {get;set;}
            //public String QuChiJieGuo { get; set; }
            //public String XueHongDanBaiJieGuo { get; set; }
            //public String TiZhongJieGuo { get; set; }
            //public String ShenGaoJieGuo { get; set; }
            //public String XueYaJieGuo { get; set; 

            //既往病史
            record.JiWangBingShi = string.Empty;
            if(record.GanYan)
            {
                record.JiWangBingShi += "肝炎;";
            }
            if(record.FeiJieHe)
            {
                record.JiWangBingShi += "肺结核;";
            }
            if(record.XinZangBing)
            {
                record.JiWangBingShi += "先天性心脏病;";
            }
            if(record.ShenYan)
            {
                record.JiWangBingShi += "肾炎;";
            }
            if(record.FengShiBing)
            {
                record.JiWangBingShi += "风湿病;";
            }
            if(!record.DiFangBing.IsNullOrEmpty())
            {
                record.JiWangBingShi += record.DiFangBing + ";";
            }
            if(!record.QiTaBing.IsNullOrEmpty())
            {
                record.JiWangBingShi += record.QiTaBing;
            }

            record.JiWangBingShi = record.JiWangBingShi.TrimEnd(';');
            if(record.JiWangBingShi.IsNullOrEmpty())
            {
                record.JiWangBingShi = "无";
            }

            //视力
            record.ShiLiJieGuo = "视力正常";
            if (record.ZuoYan < 5.0f || record.YouYan < 5.0f)
            {
                record.ShiLiJieGuo = "视力不良";
            }
            //龋齿
            record.QuChiJieGuo = record.TiJianRiQi.Year>2012 ? "无恒牙龋齿":"无龋齿";
            if (record.BZuoShang > 0 || record.BZuoXia > 0 || record.BYouShang > 0 || record.BYouXia > 0)
            {
                record.QuChiJieGuo = string.Empty;
                if (record.BZuoShang > 0)
                {
                    record.QuChiJieGuo += string.Format("{1}左上{0}B,", record.BZuoShang,record.TiJianRiQi.Year>2012 ? "恒牙":"");
                }
                if (record.BZuoXia > 0)
                {
                    record.QuChiJieGuo += string.Format("{1}左下{0}B,", record.BZuoXia, record.TiJianRiQi.Year > 2012 ? "恒牙" : "");
                }
                if (record.BYouShang > 0)
                {
                    record.QuChiJieGuo += string.Format("{1}右上{0}B,", record.BYouShang, record.TiJianRiQi.Year > 2012 ? "恒牙" : "");
                }
                if (record.BYouXia > 0)
                {
                    record.QuChiJieGuo += string.Format("{1}右下{0}B", record.BYouXia, record.TiJianRiQi.Year > 2012 ? "恒牙" : "");
                }
                record.QuChiJieGuo = record.QuChiJieGuo.TrimEnd(',');
            }


            record.StringReserved2 = "无乳牙龋齿";
            if (record.IntReserved1 > 0 || record.IntReserved2 > 0 || record.IntReserved3 > 0 || record.IntReserved4 > 0)
            {
                record.StringReserved2 = string.Empty;
                if (record.IntReserved1 > 0)
                {
                    record.StringReserved2 += string.Format("乳牙左上{0}A,", record.IntReserved1);//, record.TiJianRiQi.Year > 2012 ? "恒牙" : "");
                }
                if (record.IntReserved2 > 0)
                {
                    record.StringReserved2 += string.Format("乳牙左下{0}A,", record.IntReserved2);//, record.TiJianRiQi.Year > 2012 ? "恒牙" : "");
                }
                if (record.IntReserved3 > 0)
                {
                    record.StringReserved2 += string.Format("乳牙右上{0}A,", record.IntReserved3);//, record.TiJianRiQi.Year > 2012 ? "恒牙" : "");
                }
                if (record.IntReserved4 > 0)
                {
                    record.StringReserved2 += string.Format("乳牙右下{0}A", record.IntReserved4);//, record.TiJianRiQi.Year > 2012 ? "恒牙" : "");
                }
                record.StringReserved2 = record.StringReserved2.TrimEnd(',');
            }
            //血红蛋白结果

          //  double totalDays = (record.TiJianRiQi - record.ShengRi).TotalDays;

       //     DateTimeOffset dto = new DateTimeOffset(record.ShengRi, record.TiJianRiQi - record.ShengRi);



            int older = (int)((record.TiJianRiQi - record.ShengRi).TotalDays / 365.25f);


            record.XueHongDanBaiJieGuo = "正常";

            if((int)record.XueHongDanBai == 0)
            {
                record.XueHongDanBaiJieGuo = "未检";
            }
            else if ((older >= 5 && older <= 11 && record.XueHongDanBai < 115f) || //5~11岁，<115g/L
                (older >= 12 && older <= 14 && record.XueHongDanBai < 120f) || //12~14岁，<120g/L
                (older >= 15 && record.XingBie && record.XueHongDanBai < 130f) || //>=15岁，男，<130g/L
                (older >= 15 && !record.XingBie && record.XueHongDanBai < 120f))//>=15岁，女, <120g/L
            {
                record.XueHongDanBaiJieGuo = "贫血";
            }

            //体重,身高结果
            record.TiZhongZhiShu = (float)(record.TiZhong / Math.Pow(record.ShenGao / 100, 2));
            record.TiZhongJieGuo = Utility.GetWeightLevel(older, record.TiZhong, record.XingBie);

            record.ShenGaoJieGuo = Utility.GetHeightLevel(older, record.ShenGao, record.XingBie);

            record.StringReserved1 = "正常";//Utility.GetNutritionLevel(older, record.ShenGao, record.TiZhong, record.XingBie);
            if (older <= 6)
            {
                if (record.XingBie)
                {
                    if (record.TiZhongZhiShu > 18.1f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 16.6f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
                else
                {
                    if (record.TiZhongZhiShu > 17.9f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 16.3f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
            }
            else if (older <= 7)
            {
                if (record.XingBie)
                {
                    if (record.TiZhongZhiShu > 19.2f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 17.4f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
                else
                {
                    if (record.TiZhongZhiShu > 18.9f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 17.2f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
            }
            else if (older <= 8)
            {
                if (record.XingBie)
                {
                    if (record.TiZhongZhiShu > 20.3f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 18.1f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
                else
                {
                    if (record.TiZhongZhiShu > 19.9f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 18.1f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
            }
            else if (older <= 9)
            {
                if (record.XingBie)
                {
                    if (record.TiZhongZhiShu > 21.4f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 18.9f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
                else
                {
                    if (record.TiZhongZhiShu > 21f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 19f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
            }
            else if (older <= 10)
            {
                if (record.XingBie)
                {
                    if (record.TiZhongZhiShu > 22.5f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 19.6f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
                else
                {
                    if (record.TiZhongZhiShu > 22.1f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 20f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
            }
            else if (older <= 11)
            {
                if (record.XingBie)
                {
                    if (record.TiZhongZhiShu > 23.6f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 20.3f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
                else
                {
                    if (record.TiZhongZhiShu > 23.3f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 21.1f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
            }
            else if (older <= 12)
            {
                if (record.XingBie)
                {
                    if (record.TiZhongZhiShu > 24.7f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 21f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
                else
                {
                    if (record.TiZhongZhiShu > 24.5f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 21.9f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
            }
            else if (older <= 13)
            {
                if (record.XingBie)
                {
                    if (record.TiZhongZhiShu > 25.7f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 21.9f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
                else
                {
                    if (record.TiZhongZhiShu > 25.6f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 22.6f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
            }
            else if (older <= 14)
            {
                if (record.XingBie)
                {
                    if (record.TiZhongZhiShu > 26.4f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 22.6f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
                else
                {
                    if (record.TiZhongZhiShu > 26.3f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 23f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
            }
            else if (older <= 15)
            {
                if (record.XingBie)
                {
                    if (record.TiZhongZhiShu > 26.9f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 23.1f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
                else
                {
                    if (record.TiZhongZhiShu > 26.9f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 23.4f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
            }
            else if (older <= 16)
            {
                if (record.XingBie)
                {
                    if (record.TiZhongZhiShu > 27.4f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 23.5f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
                else
                {
                    if (record.TiZhongZhiShu > 27.4f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 23.7f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
            }
            else if (older <= 17)
            {
                if (record.XingBie)
                {
                    if (record.TiZhongZhiShu > 27.8f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 23.8f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
                else
                {
                    if (record.TiZhongZhiShu > 27.7f)
                    {
                        record.StringReserved1 = "肥胖";
                    }
                    else if (record.TiZhongZhiShu > 23.8f)
                    {
                        record.StringReserved1 = "超重";
                    }
                }
            }
            else if (older >= 18)
            {

                if (record.TiZhongZhiShu > 28f)
                {
                    record.StringReserved1 = "肥胖";
                }
                else if (record.TiZhongZhiShu > 24f)
                {
                    record.StringReserved1 = "超重";
                }

            }

            //血压
            record.XueYaJieGuo = "正常";
            if (record.SuZhangYa > 90 || record.SouSuoYa > 140)
            {
                record.XueYaJieGuo = "血压偏高";
            }
            else if (record.SuZhangYa < 60 || record.SouSuoYa < 90)
            {
                record.XueYaJieGuo = "血压偏低";
            }
        }



        private static string CreateInsertCommandText(SchoolHealthRecord record)
        {
            if (record == null) return string.Empty;

            StringBuilder sb = new StringBuilder();

            Dictionary<string, object> dicValues = new Dictionary<string, object>();

            foreach (var prop in typeof(SchoolHealthRecord).GetProperties())
            {
                if (prop.Name == "ID") continue;
                object val = prop.GetValue(record, null);
                if (val != null)
                {
                    dicValues[prop.Name] = val;
                }
            }

            if (dicValues.Count == 0) return string.Empty;

            sb.Append("Insert into SchoolHealth (");
            List<string> keys = new List<string>(dicValues.Count);
            foreach (var key in dicValues.Keys)
            {
                keys.Add(key);
            }

            for (int i = 0; i < keys.Count; i++)
            {
                if (i == keys.Count - 1)
                {
                    sb.AppendFormat("[{0}]) values (", keys[i]);
                }
                else
                {
                    sb.AppendFormat("[{0}],", keys[i]);
                }
            }

            for (int i = 0; i < keys.Count; i++)
            {
                object obj = dicValues[keys[i]];
                if (obj is string || obj is DateTime)
                {
                    if (i == keys.Count - 1)
                    {
                        sb.AppendFormat("'{0}')", obj);
                    }
                    else
                    {
                        sb.AppendFormat("'{0}',", obj);
                    }
                }
                else
                {
                    if (i == keys.Count - 1)
                    {
                        sb.AppendFormat("{0})", obj);
                    }
                    else
                    {
                        sb.AppendFormat("{0},", obj);
                    }
                }
            }

            return sb.ToString();

        }



        private static string CreateUpdateCommandText(SchoolHealthRecord record)
        {
            if (record == null) return string.Empty;

            StringBuilder sb = new StringBuilder();

            Dictionary<string, object> dicValues = new Dictionary<string, object>();

            foreach (var prop in typeof(SchoolHealthRecord).GetProperties())
            {
                if (prop.Name == "ID" || prop.Name == "GUID") continue;
                object val = prop.GetValue(record, null);
                if (val != null)
                {
                    dicValues[prop.Name] = val;
                }
            }

            if (dicValues.Count == 0) return string.Empty;

            sb.Append("Update SchoolHealth Set ");

            var ar = dicValues.ToList();

            for (int i = 0; i < ar.Count;i++ )
            {
                if (i == ar.Count - 1)
                {
                    if (ar[i].Value is string)
                    {
                        sb.AppendFormat("[{0}] = '{1}' ", ar[i].Key, ar[i].Value);
                    }
                    else
                    {
                        sb.AppendFormat("[{0}] = {1} ", ar[i].Key, ar[i].Value);
                    }
                  
                }
                else
                {
                    if (ar[i].Value is string || ar[i].Value is DateTime)
                    {
                        sb.AppendFormat("[{0}] = '{1}',", ar[i].Key,ar[i].Value);
                    }
                    else
                    {
                        sb.AppendFormat("[{0}] = {1},", ar[i].Key, ar[i].Value);
                    }
                }
            }

            sb.AppendFormat(" where [GUID] =  '{0}' ", record.GUID);

                return sb.ToString();

        }



        #region DataColumns
        private static DataGridViewTextBoxColumn CreateTextColumn(string dataProp, string header, int? width = null, bool visible = true)
        {
            DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
            col.HeaderText = header;
            col.Name = string.Format("col{0}", dataProp);
            col.DataPropertyName = dataProp;
            col.Visible = visible;
            if (width != null && width.HasValue)
            {
                col.Width = width.Value;
            }
            return col;
        }


        private static List<DataGridViewColumn> defaultColumns;
        public static List<DataGridViewColumn> DefaultColumns
        {
            get
            {
                if (defaultColumns == null)
                {
                    defaultColumns = new List<DataGridViewColumn>()
                    {
                        CreateTextColumn("ID","",null, false),
                        CreateTextColumn("XueXiao","学校"),
                        CreateTextColumn("NianJi","年级"),
                        CreateTextColumn("BanJi","班级"),
                        CreateTextColumn("XingMing","姓名"),
                          CreateTextColumn("StringReserved3","学号"),
                        CreateTextColumn("XingBie","性别"),
                        CreateTextColumn("ShengRi","出生日期"),
                        CreateTextColumn("YuHangHuJi","余杭户籍"),
                        CreateTextColumn("TiJianRiQi","体检日期"),
                        CreateTextColumn("TiJianDanWei","体检单位"),

                         CreateTextColumn("FloatReserved4","遗精/初潮年龄",120),

                        CreateTextColumn("ShenGao","身高(cm)"),
                        CreateTextColumn("TiZhong","体重(kg)"),
                          
                        CreateTextColumn("FeiHuoLiang","肺活量"),
                        CreateTextColumn("FloatReserved1","胸围(cm)"),
                                CreateTextColumn("FloatReserved2","腰围(cm)"),
                                        CreateTextColumn("FloatReserved3","臀围(cm)"),
                        CreateTextColumn("SuZhangYa","舒张压"),
                        CreateTextColumn("SouSuoYa","收缩压"),
                        CreateTextColumn("MaiBo","脉搏"),
                        CreateTextColumn("Xin",ConfigurationHelper.心),
                        CreateTextColumn("Fei",ConfigurationHelper.肺),
                        CreateTextColumn("Gan",ConfigurationHelper.肝),
                        CreateTextColumn("Pi",ConfigurationHelper.脾),
                        CreateTextColumn("ZuoYan","左裸视"),
                        CreateTextColumn("YouYan","右裸视"),
                           CreateTextColumn("ShiLiJieGuo","视力结果"),
                        CreateTextColumn("ShaYan",ConfigurationHelper.沙眼),
                        CreateTextColumn("JieMoYan",ConfigurationHelper.结膜炎),
                        CreateTextColumn("SeJue",ConfigurationHelper.色觉),
                            CreateTextColumn("QuChiJieGuo","龋齿"),
                        CreateTextColumn("YaZhou",ConfigurationHelper.牙周),
                        CreateTextColumn("Er",ConfigurationHelper.耳),
                        CreateTextColumn("Bi",ConfigurationHelper.鼻),
                        CreateTextColumn("BianTaoTi",ConfigurationHelper.扁桃体),
                        CreateTextColumn("Tou","头部"),   
                        CreateTextColumn("Jing","颈部"),
                        CreateTextColumn("Xiong","胸部"),
                        CreateTextColumn("JiZhu",ConfigurationHelper.脊柱),
                        CreateTextColumn("SiZhi",ConfigurationHelper.四肢),
                        CreateTextColumn("PiFu",ConfigurationHelper.皮肤),
                        CreateTextColumn("LinBaJie",ConfigurationHelper.淋巴结),
                        CreateTextColumn("XueXing",ConfigurationHelper.血型),
                        CreateTextColumn("XueHongDanBai","血红蛋白(g/L)"),
                         
                        CreateTextColumn("HuiChongLuan",ConfigurationHelper.蛔虫卵),
                        CreateTextColumn("KaBa",ConfigurationHelper.卡疤),
                        CreateTextColumn("JieHeJunSu",ConfigurationHelper.结核菌素试验,120),
                        CreateTextColumn("ZhuanAnMei",ConfigurationHelper.谷丙转氨酶),
                        CreateTextColumn("DanHongSu",ConfigurationHelper.胆红素),
                    CreateTextColumn("ShenGaoJieGuo","身高评价"),
                           CreateTextColumn("TiZhongJieGuo","体重评价"),
                           CreateTextColumn("StringReserved1","营养评价"),

                    };
                }
                return defaultColumns;
            }
        }
        #endregion
        static PropertyInfo[] recordProps = typeof(SchoolHealthRecord).GetProperties();

        private static SchoolHealthRecord RecordFromDataRow(DataRow dr)
        {
            SchoolHealthRecord rec = new SchoolHealthRecord();
         
            foreach (var prop in recordProps)
            {
                if (dr[prop.Name] == DBNull.Value) continue;
                prop.SetValue(rec, dr[prop.Name],null);
            }

            return rec;
        }

        private static List<SchoolHealthRecord> RecordFromDataTable(DataTable dt)
        {
            List<SchoolHealthRecord> list = new List<SchoolHealthRecord>(dt.Rows.Count);
            
            foreach (DataRow dr in dt.Rows)
            {
                list.Add(RecordFromDataRow(dr));
            }

            return list;
        }

        internal static SchoolHealthRecord GetRecord(int id)
        {
            var ds = GetByID(id);
            if(ds.Tables[0].Rows.Count > 0)
            {
                var r = RecordFromDataRow(ds.Tables[0].Rows[0]);
                return r;
            }
            else 
            {
                return null;
            }
        }

        internal static DataSet GetAll()
        {
            return SchoolHealthQueryDataSet(string.Format("select * from SchoolHealth"));
        }

        internal static DataSet GetByID(int recordId)
        {
            return SchoolHealthQueryDataSet(string.Format("select * from SchoolHealth where [ID]={0}", recordId));
        }

        internal static DataTable TiJianHuiZongTable(string school = null)
        {
            var grades = ConfigurationHelper.DataSourceByKey(ConfigurationHelper.年级);

           
            var dt = new SchoolHealth.ReportDataSet.tableTiJianHuiZongDataTable();

            

            DataSet gradeDS = null;
            DataTable tb;
            List<SchoolHealthRecord> records;
            ReportDataSet.tableTiJianHuiZongRow row;
            foreach (var grade in grades)
            {      
                if(school.IsNullOrEmpty())
                     gradeDS = SchoolHealthQueryDataSet(string.Format(@"select * from schoolhealth where [NianJi]='{0}'", grade));
                else
                    gradeDS = SchoolHealthQueryDataSet(string.Format(@"select * from schoolhealth where [NianJi]='{0}' and [XueXiao]='{1}'", grade,school));

                tb = gradeDS.Tables[0];

                if (tb.Rows.Count == 0) continue;

                records = RecordFromDataTable(tb);
                row = dt.NewtableTiJianHuiZongRow();

                row.TiJianRenShu = records.Count;
                row.BianTaoTiYan = records.Count(r => IsAbnormal(r.BianTaoTi));
                row.BiBing = records.Count(r => IsAbnormal(r.Bi));

                row.ChaoZhong = records.Count(r => r.StringReserved1 == "超重");
                row.DiXueYa = records.Count(r => r.XueYaJieGuo.Contains("低"));
                row.ErBing = records.Count(r => IsAbnormal(r.Er));
                row.FeiHuoDongXingFeiJieHe = records.Count(r => r.Fei == "非活动性肺结核");
                row.FeiPang = records.Count(r => r.StringReserved1 == "肥胖");
                row.GanZhongDa= records.Count(r => r.Gan == "肝肿大");
                row.GaoXueYa = records.Count(r => r.XueYaJieGuo.Contains("高"));
                row.HuoDongXingFeiJieHe = records.Count(r => r.Fei == "活动性肺结核");
                row.JieHeJunSuYangXing = records.Count(r => IsAbnormal(r.JieHeJunSu));
                row.JieMoYan = records.Count(r => r.JieMoYan == "结膜炎");
                row.JiZhuCeWan = records.Count(r => r.JiZhu == "脊柱侧弯");
                row.NianJi = grade;
                row.PiFuBing = records.Count(r => IsAbnormal(r.PiFu));
                row.PiZhongDa = records.Count(r => r.Pi == "脾肿大");
                row.QuChi = records.Count(r => !string.IsNullOrEmpty(r.QuChiJieGuo) && r.QuChiJieGuo != "无龋齿" && r.QuChiJieGuo != "无恒牙龋齿");
                row.RuYaQuChi = records.Count(r => !string.IsNullOrEmpty(r.StringReserved2) && r.StringReserved2 != "无龋齿" && r.StringReserved2 != "无乳牙龋齿");
                row.SeMang = records.Count(r => r.SeJue == "色盲");
                row.SeRuo = records.Count(r => r.SeJue == "色弱");
                row.ShaYan = records.Count(r =>IsAbnormal(r.ShaYan));
                row.ShiLiBuLiang = records.Count(r => r.ShiLiJieGuo.Contains("不良"));
                row.SiZhiShangCan = records.Count(r => IsAbnormal(r.SiZhi));
            //    row.TiJianRenShu = records.Count(r => r.BianTaoTi == "扁桃体炎");
                row.WuKaBa = records.Count(r => !r.KaBa);
                row.XinZangBing = records.Count(r => IsAbnormal(r.Xin));
                row.YaYinYan = records.Count(r => r.YaZhou == "牙龈炎");
                row.YuHangHuJi = records.Count(r => r.YuHangHuJi);

                row.ShenGaoShangDeng = records.Count(r => r.ShenGaoJieGuo == "上等");
                row.ShenGaoXiaDeng = records.Count(r => r.ShenGaoJieGuo == "下等");
                row.TiZhongShangDeng = records.Count(r => r.TiZhongJieGuo == "上等");
                row.TiZhongXiaDeng = records.Count(r => r.TiZhongJieGuo == "下等");




                row.YinYangBuLiang = records.Count(r =>
                    Utility.IsUnderNutritionLevel((int)((r.TiJianRiQi - r.ShengRi).TotalDays / 365.25f), r.XingBie,
                    r.TiZhong, r.ShenGao) == true);

                row.PinXue = records.Count(r => r.XueHongDanBaiJieGuo == "贫血");
                row.RuChongGanRan = records.Count(r => r.HuiChongLuan == "有虫卵");


                dt.Rows.Add(row);
            }

            
            //统计男
            if (school.IsNullOrEmpty())
                gradeDS = SchoolHealthQueryDataSet(@"select * from schoolhealth where [XingBie]=true");
            else
                gradeDS = SchoolHealthQueryDataSet(string.Format(@"select * from schoolhealth where [Xingbie]=true and [XueXiao]='{0}'", school));

            

            tb = gradeDS.Tables[0];

            if (tb.Rows.Count > 0)
            {
            //    skip++;
                records = RecordFromDataTable(tb);
                row = dt.NewtableTiJianHuiZongRow();

                row.TiJianRenShu = records.Count;
                row.BianTaoTiYan = records.Count(r => IsAbnormal(r.BianTaoTi));
                row.BiBing = records.Count(r => IsAbnormal(r.Bi));

                row.ChaoZhong = records.Count(r => r.StringReserved1 == "超重");
                row.DiXueYa = records.Count(r => r.XueYaJieGuo.Contains("低"));
                row.ErBing = records.Count(r => IsAbnormal(r.Er));
                row.FeiHuoDongXingFeiJieHe = records.Count(r => r.Fei == "非活动性肺结核");
                row.FeiPang = records.Count(r => r.StringReserved1 == "肥胖");
                row.GanZhongDa = records.Count(r => r.Gan == "肝肿大");
                row.GaoXueYa = records.Count(r => r.XueYaJieGuo.Contains("高"));
                row.HuoDongXingFeiJieHe = records.Count(r => r.Fei == "活动性肺结核");
                row.JieHeJunSuYangXing = records.Count(r => IsAbnormal(r.JieHeJunSu));
                row.JieMoYan = records.Count(r => r.JieMoYan == "结膜炎");
                row.JiZhuCeWan = records.Count(r => r.JiZhu == "脊柱侧弯");
                row.NianJi = "合计(男)";
                row.PiFuBing = records.Count(r => IsAbnormal(r.PiFu));
                row.PiZhongDa = records.Count(r => r.Pi == "脾肿大");
                row.QuChi = records.Count(r => !string.IsNullOrEmpty(r.QuChiJieGuo) && r.QuChiJieGuo != "无龋齿" && r.QuChiJieGuo != "无恒牙龋齿");
                row.RuYaQuChi = records.Count(r => !string.IsNullOrEmpty(r.StringReserved2) && r.StringReserved2 != "无龋齿" && r.StringReserved2 != "无乳牙龋齿");
                row.SeMang = records.Count(r => r.SeJue == "色盲");
                row.SeRuo = records.Count(r => r.SeJue == "色弱");
                row.ShaYan = records.Count(r => IsAbnormal(r.ShaYan));
                row.ShiLiBuLiang = records.Count(r => r.ShiLiJieGuo.Contains("不良"));
                row.SiZhiShangCan = records.Count(r => IsAbnormal(r.SiZhi));
                //    row.TiJianRenShu = records.Count(r => r.BianTaoTi == "扁桃体炎");
                row.WuKaBa = records.Count(r => !r.KaBa);
                row.XinZangBing = records.Count(r => IsAbnormal(r.Xin));
                row.YaYinYan = records.Count(r => r.YaZhou == "牙龈炎");
                row.YuHangHuJi = records.Count(r => r.YuHangHuJi);

                row.ShenGaoShangDeng = records.Count(r => r.ShenGaoJieGuo == "上等");
                row.ShenGaoXiaDeng = records.Count(r => r.ShenGaoJieGuo == "下等");
                row.TiZhongShangDeng = records.Count(r => r.TiZhongJieGuo == "上等");
                row.TiZhongXiaDeng = records.Count(r => r.TiZhongJieGuo == "下等");

                row.YinYangBuLiang = records.Count(r =>
                   Utility.IsUnderNutritionLevel((int)((r.TiJianRiQi - r.ShengRi).TotalDays / 365.25f), r.XingBie,
                   r.TiZhong, r.ShenGao) == true);
                row.PinXue = records.Count(r => r.XueHongDanBaiJieGuo == "贫血");
                row.RuChongGanRan = records.Count(r => r.HuiChongLuan == "有虫卵");
                dt.Rows.Add(row);
            }

            
            
            //统计女

            if (school.IsNullOrEmpty())
                gradeDS = SchoolHealthQueryDataSet(@"select * from schoolhealth where [XingBie]=false");
            else
                gradeDS = SchoolHealthQueryDataSet(string.Format(@"select * from schoolhealth where [Xingbie]=false and [XueXiao]='{0}'",  school));


            tb = gradeDS.Tables[0];

            if (tb.Rows.Count > 0)
            {
               // skip++;
                records = RecordFromDataTable(tb);
                row = dt.NewtableTiJianHuiZongRow();

                row.TiJianRenShu = records.Count;
                row.BianTaoTiYan = records.Count(r => IsAbnormal(r.BianTaoTi));
                row.BiBing = records.Count(r => IsAbnormal(r.Bi));

                row.ChaoZhong = records.Count(r => r.StringReserved1 == "超重");
                row.DiXueYa = records.Count(r => r.XueYaJieGuo.Contains("低"));
                row.ErBing = records.Count(r => IsAbnormal(r.Er));
                row.FeiHuoDongXingFeiJieHe = records.Count(r => r.Fei == "非活动性肺结核");
                row.FeiPang = records.Count(r => r.StringReserved1 == "肥胖");
                row.GanZhongDa = records.Count(r => r.Gan == "肝肿大");
                row.GaoXueYa = records.Count(r => r.XueYaJieGuo.Contains("高"));
                row.HuoDongXingFeiJieHe = records.Count(r => r.Fei == "活动性肺结核");
                row.JieHeJunSuYangXing = records.Count(r => IsAbnormal(r.JieHeJunSu));
                row.JieMoYan = records.Count(r => r.JieMoYan == "结膜炎");
                row.JiZhuCeWan = records.Count(r => r.JiZhu == "脊柱侧弯");
                row.NianJi = "合计(女)";
                row.PiFuBing = records.Count(r => IsAbnormal(r.PiFu));
                row.PiZhongDa = records.Count(r => r.Pi == "脾肿大");
                row.QuChi = records.Count(r => !string.IsNullOrEmpty(r.QuChiJieGuo) && r.QuChiJieGuo != "无龋齿" && r.QuChiJieGuo != "无恒牙龋齿");
                row.RuYaQuChi = records.Count(r => !string.IsNullOrEmpty(r.StringReserved2) && r.StringReserved2 != "无龋齿" && r.StringReserved2 != "无乳牙龋齿");
                row.SeMang = records.Count(r => r.SeJue == "色盲");
                row.SeRuo = records.Count(r => r.SeJue == "色弱");
                row.ShaYan = records.Count(r => IsAbnormal(r.ShaYan));
                row.ShiLiBuLiang = records.Count(r => r.ShiLiJieGuo.Contains("不良"));
                row.SiZhiShangCan = records.Count(r => IsAbnormal(r.SiZhi));
                //    row.TiJianRenShu = records.Count(r => r.BianTaoTi == "扁桃体炎");
                row.WuKaBa = records.Count(r => !r.KaBa);
                row.XinZangBing = records.Count(r => IsAbnormal(r.Xin));
                row.YaYinYan = records.Count(r => r.YaZhou == "牙龈炎");
                row.YuHangHuJi = records.Count(r => r.YuHangHuJi);

                row.ShenGaoShangDeng = records.Count(r => r.ShenGaoJieGuo == "上等");
                row.ShenGaoXiaDeng = records.Count(r => r.ShenGaoJieGuo == "下等");
                row.TiZhongShangDeng = records.Count(r => r.TiZhongJieGuo == "上等");
                row.TiZhongXiaDeng = records.Count(r => r.TiZhongJieGuo == "下等");

                row.YinYangBuLiang = records.Count(r =>
                   Utility.IsUnderNutritionLevel((int)((r.TiJianRiQi - r.ShengRi).TotalDays / 365.25f), r.XingBie,
                   r.TiZhong, r.ShenGao) == true);
                row.PinXue = records.Count(r => r.XueHongDanBaiJieGuo == "贫血");
                row.RuChongGanRan = records.Count(r => r.HuiChongLuan == "有虫卵");
                dt.Rows.Add(row);
            }

            //合计

            if (school.IsNullOrEmpty())
                gradeDS = SchoolHealthQueryDataSet(@"select * from schoolhealth");
            else
                gradeDS = SchoolHealthQueryDataSet(string.Format(@"select * from schoolhealth where [XueXiao]='{0}'", school));


            tb = gradeDS.Tables[0];

            if (tb.Rows.Count > 0)
            {
                
                records = RecordFromDataTable(tb);
                row = dt.NewtableTiJianHuiZongRow();

                row.TiJianRenShu = records.Count;
                row.BianTaoTiYan = records.Count(r => IsAbnormal(r.BianTaoTi));
                row.BiBing = records.Count(r => IsAbnormal(r.Bi));

                row.ChaoZhong = records.Count(r => r.StringReserved1 == "超重");
                row.DiXueYa = records.Count(r => r.XueYaJieGuo.Contains("低"));
                row.ErBing = records.Count(r => IsAbnormal(r.Er));
                row.FeiHuoDongXingFeiJieHe = records.Count(r => r.Fei == "非活动性肺结核");
                row.FeiPang = records.Count(r => r.StringReserved1 == "肥胖");
                row.GanZhongDa = records.Count(r => r.Gan == "肝肿大");
                row.GaoXueYa = records.Count(r => r.XueYaJieGuo.Contains("高"));
                row.HuoDongXingFeiJieHe = records.Count(r => r.Fei == "活动性肺结核");
                row.JieHeJunSuYangXing = records.Count(r => IsAbnormal(r.JieHeJunSu));
                row.JieMoYan = records.Count(r => r.JieMoYan == "结膜炎");
                row.JiZhuCeWan = records.Count(r => r.JiZhu == "脊柱侧弯");
                row.NianJi = "合计";
                row.PiFuBing = records.Count(r => IsAbnormal(r.PiFu));
                row.PiZhongDa = records.Count(r => r.Pi == "脾肿大");
                row.QuChi = records.Count(r => !string.IsNullOrEmpty(r.QuChiJieGuo) && r.QuChiJieGuo != "无龋齿" && r.QuChiJieGuo != "无恒牙龋齿");
                row.RuYaQuChi = records.Count(r => !string.IsNullOrEmpty(r.StringReserved2) && r.StringReserved2 != "无龋齿" && r.StringReserved2 != "无乳牙龋齿");
                row.SeMang = records.Count(r => r.SeJue == "色盲");
                row.SeRuo = records.Count(r => r.SeJue == "色弱");
                row.ShaYan = records.Count(r => IsAbnormal(r.ShaYan));
                row.ShiLiBuLiang = records.Count(r => r.ShiLiJieGuo.Contains("不良"));
                row.SiZhiShangCan = records.Count(r => IsAbnormal(r.SiZhi));
                //    row.TiJianRenShu = records.Count(r => r.BianTaoTi == "扁桃体炎");
                row.WuKaBa = records.Count(r => !r.KaBa);
                row.XinZangBing = records.Count(r => IsAbnormal(r.Xin));
                row.YaYinYan = records.Count(r => r.YaZhou == "牙龈炎");
                row.YuHangHuJi = records.Count(r => r.YuHangHuJi);

                row.ShenGaoShangDeng = records.Count(r => r.ShenGaoJieGuo == "上等");
                row.ShenGaoXiaDeng = records.Count(r => r.ShenGaoJieGuo == "下等");
                row.TiZhongShangDeng = records.Count(r => r.TiZhongJieGuo == "上等");
                row.TiZhongXiaDeng = records.Count(r => r.TiZhongJieGuo == "下等");


                row.YinYangBuLiang = records.Count(r =>
                   Utility.IsUnderNutritionLevel((int)((r.TiJianRiQi - r.ShengRi).TotalDays / 365.25f), r.XingBie,
                   r.TiZhong, r.ShenGao) == true);
                row.PinXue = records.Count(r => r.XueHongDanBaiJieGuo == "贫血");
                row.RuChongGanRan = records.Count(r => r.HuiChongLuan == "有虫卵");
                dt.Rows.Add(row);
            }


                return dt;
        }

        private static bool IsAbnormal(string str)
        {
            return !(str.IsNullOrEmpty() || str == "正常" || str == "未检" || str == "阴性" || str == "标准");
        }

        internal static DataTable TiJianJieGuoTable(DataTable table = null,string school = null)
        {
            if(table == null)
            {
                if (school.IsNullOrEmpty())
                    table = GetAll().Tables[0];
                else
                    table = SchoolHealthQueryDataSet(string.Format(@"select * from schoolhealth where [XueXiao]='{0}'", school)).Tables[0];
            }
                 var dt = new SchoolHealth.ReportDataSet.tableTiJianJieGuoDataTable();
                 var r = dt.NewtableTiJianJieGuoRow();
                 var props = r.GetType().GetProperties();
                 Dictionary<string, PropertyInfo> dicColProp = new Dictionary<string, PropertyInfo>();

                 foreach (var prop in props)
                 {
                     if (table.Columns.Contains(prop.Name))
                     {
                         dicColProp.Add(prop.Name,prop);
                     }
                 }
   
            foreach (DataRow dr  in table.Rows)
            {
                r = dt.NewtableTiJianJieGuoRow();
               
                foreach (var kp in dicColProp)
                {
                    if (dr[kp.Key] == DBNull.Value) continue;

                    kp.Value.SetValue(r, dr[kp.Key], null);
                }

                dt.Rows.Add(r);
            }

        
           
                return dt;
        
          
        }

        internal static DataTable FanKuiDanDataTable(int recordId, List<ReportParameter> rptParameters)
        {
            var reader = GetByID(recordId);
            var dt = new SchoolHealth.ReportDataSet.tableFanKuiDanDataTable();
            //dt.Columns.Add(new DataColumn("XueXiao"));
            //dt.Columns.Add(new DataColumn("NianJi"));
            //dt.Columns.Add(new DataColumn("BanJi"));
            //dt.Columns.Add(new DataColumn("XingMing"));
            //dt.Columns.Add(new DataColumn("XingBie", typeof(bool)));
            //dt.Columns.Add(new DataColumn("ShengRi", typeof(DateTime)));
            //dt.Columns.Add(new DataColumn("ShenGao", typeof(float)));
            //dt.Columns.Add(new DataColumn("TiZhong", typeof(float)));
            //dt.Columns.Add(new DataColumn("TiZhongZhiShu", typeof(float)));
            //dt.Columns.Add(new DataColumn("TiZhongJieGuo"));
            //dt.Columns.Add(new DataColumn("QuChiJieGuo"));
            ////dt.Columns.Add(new DataColumn("TiJianRiQi",typ));
            ////dt.Columns.Add(new DataColumn("TiJianDanWei"));
            //dt.Columns.Add(new DataColumn("ZuoYan", typeof(float)));
            //dt.Columns.Add(new DataColumn("ShiLiJieGuo"));
            //dt.Columns.Add(new DataColumn("YouYan", typeof(float)));


            DataTable table = reader.Tables[0];
            if (table.Rows.Count > 0)
            {
                DataRow firstRow = table.Rows[0];
                DataRow dr = dt.NewRow();

                string hyqc = firstRow["QuChiJieGuo"].ToString().Replace("B", "个恒牙龋齿");
                string ryqc = firstRow["StringReserved2"].ToString().Replace("A", "个乳牙龋齿");

                string qc = hyqc + ";" + ryqc;


                dr.ItemArray = new object[]
                {
                    firstRow["XueXiao"],firstRow["NianJi"],firstRow["BanJi"],firstRow["XingMing"],firstRow["XingBie"],
                    firstRow["ShengRi"],firstRow["ShenGao"],firstRow["TiZhong"],firstRow["TiZhongZhiShu"],firstRow["TiZhongJieGuo"],
                   qc,firstRow["ZuoYan"],firstRow["ShiLiJieGuo"],firstRow["YouYan"],firstRow["StringReserved1"],
                    firstRow["SuZhangYa"],firstRow["SouSuoYa"],firstRow["XueYaJieGuo"],
                };
                dt.Rows.Add(dr);


                StringBuilder sb = new StringBuilder();
                string qita = string.Empty;
                //AppendAbnormal(sb, firstRow, "XueYaJieGuo", "血压");
                AppendAbnormal(sb, firstRow, "Xin", "心脏");
                AppendAbnormal(sb, firstRow, "Fei", "肺部");
                AppendAbnormal(sb, firstRow, "Gan", "肝脏");
                AppendAbnormal(sb, firstRow, "Pi", "脾脏");
                AppendAbnormal(sb, firstRow, "ShaYan", "沙眼");
                AppendAbnormal(sb, firstRow, "JieMoYan", "结膜炎");
                AppendAbnormal(sb, firstRow, "SeJue", "色觉");
                AppendAbnormal(sb, firstRow, "YaZhou", "牙周");
                AppendAbnormal(sb, firstRow, "Er", "耳");
                AppendAbnormal(sb, firstRow, "Bi", "鼻");
                AppendAbnormal(sb, firstRow, "BianTaoTi", "扁桃体");
                AppendAbnormal(sb, firstRow, "Tou", "头部");
                AppendAbnormal(sb, firstRow, "Jing", "颈部");
                AppendAbnormal(sb, firstRow, "Xiong", "胸部");
                AppendAbnormal(sb, firstRow, "SiZhi", "四肢");
                AppendAbnormal(sb, firstRow, "LinBaJie", "淋巴结");
                AppendAbnormal(sb, firstRow, "KaBa", "卡疤", "有");
                AppendAbnormal(sb, firstRow, "HuiChongLuan", "蛔虫卵", "无虫卵");
                AppendAbnormal(sb, firstRow, "JieHeJunSu", "结核菌素", "阴性");
                AppendAbnormal(sb, firstRow, "ZhuanAnMei", "谷丙转氨酶");
                AppendAbnormal(sb, firstRow, "DanHongSu", "胆红素");
                AppendAbnormal(sb, firstRow, "XueHongDanBaiJieGuo", "营养状况");
                qita = sb.ToString().Trim(';');
                if(qita.IsNullOrEmpty())
                {
                    qita = "无";
                }
                rptParameters.Add(new ReportParameter("ReportParameter_QiTaWenTi", qita));


                string result = firstRow["StringReserved1"].ToString();
                rptParameters.Add(new ReportParameter("ReportParameter_Check1", (result.Contains("超重") || result.Contains("肥胖")).ToString()));

                result = firstRow["ShiLiJieGuo"].ToString();
                rptParameters.Add(new ReportParameter("ReportParameter_Check2", (result.Contains("视力不良")).ToString()));

                if (((byte)firstRow["BZuoShang"]) > 0 || ((byte)firstRow["BZuoXia"]) > 0 || ((byte)firstRow["BYouShang"]) > 0 || ((byte)firstRow["BYouXia"]) > 0)
                {
                    rptParameters.Add(new ReportParameter("ReportParameter_Check3", true.ToString()));
                }
                else 
                {
                    rptParameters.Add(new ReportParameter("ReportParameter_Check3", false.ToString()));
                }

                result = firstRow["ShaYan"].ToString();

                rptParameters.Add(new ReportParameter("ReportParameter_Check4", (result.Contains("沙眼") || result.Contains("疑沙")).ToString()));


                rptParameters.Add(new ReportParameter("ReportParameter_Check5", (qita != "无").ToString()));

                rptParameters.Add(new ReportParameter("ReportParameter_Checker", firstRow["TiJianDanWei"].ToString()));

                DateTime date = (DateTime)firstRow["TiJianRiQi"];

                rptParameters.Add(new ReportParameter("ReportParameter_CheckDate", date.ToShortDateString()));

            }


            return dt;

        }

        private static void AppendAbnormal(StringBuilder sb, DataRow dr, string field, string title, string normal = "正常", string undo = "未检")
        {
            if (field != "KaBa")
            {
                string tmp = dr[field].ToString();
                if (!(tmp.IsNullOrEmpty() || tmp == normal || tmp == undo))
                {
                    sb.AppendFormat("{0}：{1};", title, tmp);
                }
            }
            else
            {
                var tmp = (bool)dr[field];
                if (!tmp)
                {
                    sb.AppendFormat("{0}：{1};", title, "无");
                }
            }

        }

        internal static void Delete(int id)
        {
            SchoolHealthNonQueryCommand(string.Format(@"delete from schoolhealth where [ID]={0}", id));
        }

        internal static bool Update(SchoolHealthRecord record)
        {
            UpdateStaticResult(record);

                string strSql = CreateUpdateCommandText(record);
            return SchoolHealthNonQueryCommand(strSql);
        }

        internal static bool Backup(string backupPath)
        {
            if(conn.State == ConnectionState.Open)
            {
                conn.Close();
            }

            string src = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"Data.mdb");

            try
            {
                File.Copy(src, backupPath);
                return true;
            }
            catch 
            {
                return false;
            }
        

        }

        internal static void AutoBackup()
        {
            string src = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Backup");
            if(!Directory.Exists(src))
            {
                try
                {
                    Directory.CreateDirectory(src);
                }
                catch
                {
                	return;
                }
            }
                      string savePath = Path.Combine(src, string.Format("{0:yyyy年MM月dd日HH时mm分ss秒}备份.mdb", DateTime.Now));
                      Backup(savePath);
        }

        internal static void Restore(string src)
        {
            string newName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), Guid.NewGuid().ToString() + ".mdb");

        
            File.Copy(src, newName);


            using(OleDbConnection myConn = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}",Path.GetFileName(newName))))
            {
                using(ProgressForm pf = new ProgressForm())
                {
                    pf.SupportCancel = true;
                    pf.IsReportProgress = true;
                    pf.Text = "正在导入数据";
                    int rowCount = 0;
                    int i = 0;
                    pf.DoWork += (s, e) =>
                        {
                            try
                            {
                                myConn.Open();
                                var ds = new DataSet();
                                string strSql = @"select * from schoolhealth";
                                new OleDbDataAdapter(strSql, myConn).Fill(ds, "SchoolHealth");
                                rowCount = ds.Tables[0].Rows.Count;
                                string strGUID = @"select [id] from schoolhealth where GUID='{0}'";
                                foreach (DataRow dr in ds.Tables[0].Rows)
                                {
                                    if (pf.CancellationPending)
                                    {
                                        break;
                                    }

                                    SchoolHealthRecord r = RecordFromDataRow(dr);

                                    var sql = string.Format(strGUID, r.GUID.ToString());
                                    var reader = SchoolHealthDataReader(sql);

                                    if (reader.HasRows)//exist,then update
                                    {
                                        Update(r);
                                    }
                                    else
                                    {
                                        r.ID = 0;
                                        Add(r, false);
                                    }

                                    i++;
                                    pf.ReportProgress((int)(i * 100f / rowCount), null);
                                }
                            }
                            catch(Exception ex)
                            {
                                Utility.ShowInformation("导入发生错误:\r\n"+ex.Message);
                            }
                        };

                    pf.ProgressChanged += (s, e) =>
                        {
                            pf.Text = string.Format("正在导入{0}/{1}", i, rowCount);
                        };

                    pf.RunWorkerCompleted += (s, e) =>
                        {
                            if(myConn != null && myConn.State != ConnectionState.Closed)
                            {
                                myConn.Close();
                            }
                            if(File.Exists(newName))
                            {
                                try
                                {
                                    File.Delete(newName);
                                }
                                catch 
                                {
                                    
                                }
                            }
                        };

                    pf.Start();
                }
              
            }
        }

        internal static void UpdateStatisticResult(bool showConfirm = false)
        {
            var ds = GetAll();
            int rowCount = ds.Tables[0].Rows.Count;

            if (rowCount == 0) return;

            if(showConfirm)
                Utility.ShowInformation("统计结果需要更新,请按确定继续");

                using (ProgressForm pf = new ProgressForm())
                {
                    pf.SupportCancel = false;
                    pf.IsReportProgress = true;
                    pf.Text = "正在更新统计结果...";
           
                    int i = 0;
                    pf.DoWork += (s, e) =>
                    {
                      
                        foreach (DataRow dr in ds.Tables[0].Rows)
                        {
                            if (pf.CancellationPending)
                            {
                                break;
                            }

                            SchoolHealthRecord r = RecordFromDataRow(dr);                   
                            Update(r);
                         
                            i++;
                            pf.ReportProgress((int)(i * 100f / rowCount), null);
                        }
                    };

                    pf.RunWorkerCompleted += (s, e) =>
                        {  
                            if(!e.Cancelled)
                            {
                                //flag == 1, ver 1002
                                //flag == 2, ver 1003
                                Settings.Default.UpdateFlag = 2;
                                Settings.Default.Save();
                            }
                        
                        };

                    pf.ProgressChanged += (s, e) =>
                    {
                        pf.Text = string.Format("正在更新统计结果{0}/{1}", i, rowCount);
                    };

                    pf.Start();
                }

         
        }

        //internal static void CheckUpdate()
        //{

        //    var ds = SchoolHealthQueryDataSet(@"select * from schoolhealth where IntReserved1 <> 1");
        //        int    rowCount = ds.Tables[0].Rows.Count;

        //    if(rowCount > 0)
        //    {

        //        Utility.SaveInvoke(Form.ActiveForm,()=>
        //            {
        //                Utility.ShowInformation("发现统计结果需要更新,请按确定继续!");


        //                using (ProgressForm pf = new ProgressForm())
        //                {
        //                    pf.SupportCancel = true;
        //                    pf.IsReportProgress = true;
        //                    pf.Text = "正在更新统计结果...";

        //                    int i = 0;
        //                    pf.DoWork += (s, e) =>
        //                    {

        //                        foreach (DataRow dr in ds.Tables[0].Rows)
        //                        {
        //                            SchoolHealthRecord r = RecordFromDataRow(dr);
        //                            Update(r);

        //                            i++;
        //                            pf.ReportProgress((int)(i * 100f / rowCount), null);
        //                        }
        //                    };

        //                    pf.ProgressChanged += (s, e) =>
        //                    {
        //                        pf.Text = string.Format("正在更新统计结果{0}/{1}", i, rowCount);
        //                    };

        //                    pf.Start();
        //                }
        //            },false,false);
        //    }
        //}
    }
}
