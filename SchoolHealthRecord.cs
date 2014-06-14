using System;
using System.Collections.Generic;
using System.Text;

namespace SchoolHealth
{
    public class SchoolHealthRecord
    {
        public SchoolHealthRecord()
        {
            XingBie = true;
            YuHangHuJi = true;
            ShengRi = new DateTime(1990, 1, 1);
            ShengBingRiQi = new DateTime(2000, 1, 1);
            TiJianRiQi = new DateTime(DateTime.Now.Year,9,1);
            IntReserved1 = IntReserved2 = IntReserved3 = IntReserved4 = 0;
            FloatReserved2 = FloatReserved3 = 0;
        }

        //数据库ID
        public int ID { get; set; }
        public String GUID { get; set; }

        //学校+学生+体检基本信息
        public String XueXiao { get; set; }
        public String NianJi { get; set; }
        public String BanJi { get; set; }
        public String XingMing { get; set; }
        public bool XingBie { get; set; }
        public DateTime ShengRi { get; set; }
        public bool YuHangHuJi { get; set; }
        public DateTime TiJianRiQi { get; set; }
        public String TiJianDanWei { get; set; }

        //既往病史
        public bool GanYan { get; set; }
        public bool FeiJieHe { get; set; }
        public bool XinZangBing { get; set; }
        public bool ShenYan { get; set; }
        public bool FengShiBing { get; set; }
        public String DiFangBing { get; set; }
        public String QiTaBing { get; set; }
        public DateTime ShengBingRiQi { get; set; }

        //形体机能
        public float ShenGao { get; set; }
        public float TiZhong { get; set; }
        public int FeiHuoLiang { get; set; }
        public int SuZhangYa { get; set; }
        public int SouSuoYa { get; set; }
        public int MaiBo { get; set; }

        //内科
        public String Xin { get; set; }
        public String Fei { get; set; }
        public String Gan { get; set; }
        public String Pi { get; set; }

       //眼科
        public float ZuoYan { get; set; }
        public float YouYan { get; set; }
        public String ShaYan { get; set; }
        public String JieMoYan { get; set; }
        public String SeJue { get; set; }

        //口腔科

        public int IntReserved1 {get;set;} //A 左上
        public int IntReserved2 { get; set; } //A 左下
        public int IntReserved3 { get; set; } //A 右上
        public int IntReserved4 { get; set; } //A 右下


        public byte BZuoShang { get; set; }
        public byte BZuoXia { get; set; }
        public byte BYouShang { get; set; }
        public byte BYouXia { get; set; }
        public byte DZuoShang { get; set; }
        public byte DZuoXia { get; set; }
        public byte DYouShang { get; set; }
        public byte DYouXia { get; set; }
        public byte FZuoShang { get; set; }
        public byte FZuoXia { get; set; }
        public byte FYouShang { get; set; }
        public byte FYouXia { get; set; }
        public String YaZhou { get; set; }

        //耳鼻咽喉科
        public String Er { get; set; }
        public String Bi { get; set; }
        public String BianTaoTi { get; set; }

        //外科
        public String Tou { get; set; }
        public String Jing { get; set; }
        public String Xiong { get; set; }
        public String JiZhu { get; set; }
        public String SiZhi { get; set; }
        public String PiFu { get; set; }
        public String LinBaJie { get; set; }

        //化验
        public String XueXing { get; set; }
        public float XueHongDanBai { get; set; }
        public String HuiChongLuan { get; set; }

        //结核病
        public bool KaBa { get; set; }
        public String JieHeJunSu { get; set; }

        //肝功能
        public String ZhuanAnMei { get; set; }
        public String DanHongSu { get; set; }

        //统计结果
        public String ShiLiJieGuo {get;set;}
        //恒牙龋齿结果    
        public String QuChiJieGuo { get; set; }
        public String XueHongDanBaiJieGuo { get; set; }
        //体重上下等评价
        public String TiZhongJieGuo { get; set; }
        public String ShenGaoJieGuo { get; set; }
        public String XueYaJieGuo { get; set; }
        public float TiZhongZhiShu {get;set;}
        public string JiWangBingShi {get;set;}

        //营养评价(体重指数评价,BMI)
        public string StringReserved1 {get;set;}

        //胸围(cm)
        public float FloatReserved1 { get; set; }

        //乳牙龋齿结果
        public string StringReserved2 { get; set; }

        //腰围(cm)
        public float FloatReserved2 { get; set; }

        //臀围(cm)
        public float FloatReserved3 { get; set; }

        //遗精/初潮年龄
        public float FloatReserved4 { get; set; }

        //学号
        public string StringReserved3 { get; set; }
    }


    
 
}
