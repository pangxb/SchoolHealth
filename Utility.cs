using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using System.IO;


using System.Reflection;
using Microsoft.Reporting.WinForms;
using System.Threading;
using SchoolHealth.Properties;
using System.Management;

using System.Drawing.Printing;
using System.Drawing.Imaging;
using System.Diagnostics;


namespace SchoolHealth
{
    internal static class Utility
    {
        private static string defaultPrinter;
        private static int pageIndex;
        private static List<Stream> reportStreams;
        private static BindingFlags flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

        //<年龄<身高范围>>
        static Dictionary<int, List<float>> dicHB = new Dictionary<int, List<float>>();
        static Dictionary<int, List<float>> dicHG = new Dictionary<int, List<float>>();
        //<年龄<体重范围>>
        static Dictionary<int, List<float>> dicWB = new Dictionary<int, List<float>>();
        static Dictionary<int, List<float>> dicWG = new Dictionary<int, List<float>>();

        //<年龄<身高<身高范围>>>
        static Dictionary<int, Dictionary<int, List<float>>> dicNutBoy = new Dictionary<int, Dictionary<int, List<float>>>();
        static Dictionary<int, Dictionary<int, List<float>>> dicNutGirl = new Dictionary<int, Dictionary<int, List<float>>>();


        //<年龄<体重<身高范围>>>
        static Dictionary<int, Dictionary<float, List<float>>> dicNutBoyEx = new Dictionary<int, Dictionary<float, List<float>>> ();
        static Dictionary<int, Dictionary<float, List<float>>> dicNutGirlEx = new Dictionary<int, Dictionary<float, List<float>>> ();



        static Utility()
        {
            defaultPrinter = Settings.Default.DefaultPrinter;
            ReadStandard();
         }

        public static RibbonPanel NewRibbonPanel {get;set;}

        internal static bool SaveInvoke(Control ctrl, MethodInvoker mi, bool async, bool thread)
        {

            if (ctrl == null || mi == null)
            {
                return false;
            }

            try
            {

                //double checking
                if (ctrl.IsHandleCreated && !ctrl.Disposing && !ctrl.IsDisposed)
                {
                    if (thread)
                    {
                        ThreadPool.QueueUserWorkItem(s =>
                        {
                            if (ctrl.IsHandleCreated && !ctrl.Disposing && !ctrl.IsDisposed)
                            {
                                if (async)
                                    ctrl.BeginInvoke(mi);
                                else
                                    ctrl.Invoke(mi);
                            }
                        });
                        return true;
                    }
                    else
                    {
                        if (async)
                        {
                            if (ctrl.IsHandleCreated && !ctrl.Disposing && !ctrl.IsDisposed)
                                ctrl.BeginInvoke(mi);
                        }
                        else
                        {
                            if (ctrl.IsHandleCreated && !ctrl.Disposing && !ctrl.IsDisposed)
                                ctrl.Invoke(mi);
                        }
                        return true;
                    }
                }
                else
                {
                    return false;
                }
                //}
            }
            catch (Exception ex)
            {
            
                return false;
            }


        }


        internal static DialogResult ShowQuestion(string msg, IWin32Window owner = null, bool topMost = false)
        {
           if(owner != null)
           {
               return MessageBoxEx.Show(owner, msg, "确认信息", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
           }
           else 
           {
               return MessageBoxEx.Show(msg, "确认信息", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
           }
        }

        internal static  DialogResult ShowInformation(string msg, IWin32Window owner = null, bool topMost = false)
        {
            if (owner != null)
            {
                return MessageBoxEx.Show(owner, msg, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                return MessageBoxEx.Show(msg, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        internal static void PrintFanKuiDan(ReportViewer rptView, int recordId)
        {
            try
            {

                //rptView.LocalReport.ReportEmbeddedResource = null;
                //rptView.RefreshReport();
                //rptView.LocalReport.ReportEmbeddedResource = string.Format("SchoolHealth.反馈单.rdlc");

                rptView.LocalReport.DataSources.Clear();
                List<ReportParameter> rptParameters = new List<ReportParameter>();
                System.Data.DataTable singleRecordTable = DBHelper.FanKuiDanDataTable(recordId, rptParameters);
                var dataSet = new ReportDataSource("DataSetFanKuiDan", singleRecordTable);
                rptView.LocalReport.DataSources.Add(dataSet);
                rptView.LocalReport.SetParameters(rptParameters.ToArray());

                rptView.RefreshReport();
                Utility.PrintReport(rptView, rptView, false);
            }
            catch (System.Exception ex)
            {
                ShowInformation("发生错误:\r\n" + ex.Message);
            }

        }

        internal static void ReportFanKuiDan(ReportViewer rptView,int recordId)
        {
            try
            {

                //rptView.LocalReport.ReportEmbeddedResource = null;
                //rptView.RefreshReport();
                //rptView.LocalReport.ReportEmbeddedResource = string.Format("SchoolHealth.反馈单.rdlc");
               
                rptView.LocalReport.DataSources.Clear();
                List<ReportParameter> rptParameters = new List<ReportParameter>();
                System.Data.DataTable singleRecordTable = DBHelper.FanKuiDanDataTable(recordId, rptParameters);
                var dataSet = new ReportDataSource("DataSetFanKuiDan", singleRecordTable);
                rptView.LocalReport.DataSources.Add(dataSet);
                rptView.LocalReport.SetParameters(rptParameters.ToArray());
            
                rptView.RefreshReport();
            }
            catch (System.Exception ex)
            {
                ShowInformation("发生错误:\r\n" + ex.Message);
            }
      
        }

        internal static void ReportTiJianJieGuo(ReportViewer rptView,System.Data.DataTable dt,string school = null)
        {
            try
            {
                //rptView.LocalReport.ReportEmbeddedResource = null;
                //rptView.RefreshReport();
                //rptView.LocalReport.ReportEmbeddedResource = string.Format("SchoolHealth.体检结果.rdlc");
                rptView.LocalReport.DataSources.Clear();

                var dataSet = new ReportDataSource("DataSetTiJianJieGuo", dt);
                rptView.LocalReport.DataSources.Add(dataSet);

                var s = school ?? Properties.Settings.Default.DefaultSchool;
                if (s == null) s = string.Empty;


                if (s.IsNullOrEmpty())
                {

                    s = "体检结果";
                }
                else
                {
                    s += "体检结果";
                }

                ReportParameter rp = new ReportParameter("ReportParameter_Title",s );
                rptView.LocalReport.SetParameters(new ReportParameter[] { rp });
                rptView.RefreshReport();
            }
            catch (System.Exception ex)
            {
                ShowInformation("发生错误:\r\n" + ex.Message);
            }
        
        }

        internal static void ReportTiJianHuiZong(ReportViewer rptView,string school = null)
        {
            try
            {
                //rptView.LocalReport.ReportEmbeddedResource = null;
                //rptView.RefreshReport();
                //rptView.LocalReport.ReportEmbeddedResource = string.Format("SchoolHealth.体检汇总.rdlc");
                rptView.LocalReport.DataSources.Clear();
                var dt = DBHelper.TiJianHuiZongTable(school);
                var dataSet = new ReportDataSource("DataSetTiJianHuiZong", dt);
                rptView.LocalReport.DataSources.Add(dataSet);

                string s = school ?? Properties.Settings.Default.DefaultSchool;

                if(s.IsNullOrEmpty())
                {
                   
                    s = "体检汇总";
                }
                else 
                {
                    s += "体检汇总";
                }


                ReportParameter rp = new ReportParameter("ReportParameter_Title", s);
                rptView.LocalReport.SetParameters(new ReportParameter[] { rp });
                rptView.RefreshReport();
            }
            catch (System.Exception ex)
            {
                ShowInformation("发生错误:\r\n" + ex.Message);
            }

        }

        internal static string DefaultPrinter
        {
            set
            {
                defaultPrinter = value;
                Settings.Default.DefaultPrinter = defaultPrinter;
                Settings.Default.Save();
            }
        }

        internal static List<string> LoadPrinters()
        {
            List<string> printers = new List<string>();

            try
            {
                string wmiSQL = "SELECT DeviceId FROM Win32_Printer";
                ManagementObjectCollection printers1 = new ManagementObjectSearcher(wmiSQL).Get();

                foreach (ManagementObject printer in printers1)
                {
                    PropertyDataCollection.PropertyDataEnumerator pde = printer.Properties.GetEnumerator();

                    while (pde.MoveNext())
                    {
                        printers.Add(pde.Current.Value.ToString());
                    }
                }

                if (printers.Count > 0)
                {
                    string defaultPrintName = Settings.Default.DefaultPrinter;
                    if (string.IsNullOrEmpty(defaultPrintName) || !printers.Contains(defaultPrintName))
                    {
                        DefaultPrinter = printers[0];
                    }
                }
            }
            catch
            {
            	
            }
            finally
            {
            }
            
            return printers;
        }

        internal static bool ZoomReport(ReportViewer viewer, int index, int percent)
        {
            if (index == 0)
            {
                viewer.ZoomMode = ZoomMode.PageWidth;
            }
            else if (index == 1)
            {
                viewer.ZoomMode = ZoomMode.FullPage;
            }
            else
            {
                viewer.ZoomMode = ZoomMode.Percent;
                viewer.ZoomPercent = percent;
            }
            return true;
        }
        /// <summary>
        /// Exports the report.
        /// </summary>
        /// <param name="viewer">The viewer.</param>
        /// <param name="index">
        /// The index.
        /// 0 ---> doc
        /// 1 ---> xls
        /// 2 ---> pdf
        /// </param>
        /// <returns></returns>
        /// 
     
        internal static bool ExportReport(ReportViewer viewer, ExportType index)
        {
            if (viewer.LocalReport.DataSources.Count == 0)
            {
                return false;
            }
            Type t = viewer.GetType();
            RenderingExtension re = null;
            RenderingExtension[] extensionArray = viewer.LocalReport.ListRenderingExtensions();

            try
            {
                re = extensionArray[(int)index];

                if (re != null)
                {
                    ReportExportEventArgs e = new ReportExportEventArgs(re);
                    string methodName = "OnExport";
                    t.InvokeMember(methodName, flags | BindingFlags.InvokeMethod, null, viewer, new object[] { null, e });
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch
            {
                return false;
            }

        }

        internal static bool ReportPageSetup(ReportViewer viewer)
        {

            Type t = viewer.GetType();


            try
            {
                string methodName = "OnPageSetup";
                t.InvokeMember(methodName, flags | BindingFlags.InvokeMethod, null, viewer, new object[] { null, EventArgs.Empty });
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        ///report prints the preview .
        /// </summary>
        /// <param name="viewer">The viewer.</param>
        /// <param name="mode">The mode.</param>
        /// <returns></returns>
        internal static bool PrintPreviewReport(ReportViewer viewer, DisplayMode mode)
        {
            viewer.SetDisplayMode(mode);
            return true;
        }

        /// <summary>
        /// Prints the report.
        /// </summary>
        /// <param name="viewer">The viewer.</param>
        /// <returns></returns>
        internal static bool PrintReport(ReportViewer viewer, Control host, bool needRefresh = true)
        {
            if (!host.IsHandleCreated)
            {
                host.CreateControl();
            }
            if (string.IsNullOrEmpty(defaultPrinter))
            {
                Utility.ShowInformation("请先设置默认打印机[系统设置->打印机设置]");
                return false;
            }
            if (reportStreams != null && reportStreams.Count > 0)
            {
                foreach (Stream s in reportStreams)
                {
                    s.Dispose();
                }
                reportStreams.Clear();
            }
            else
            {
                reportStreams = new List<Stream>();
            }

            if(needRefresh)
                viewer.RefreshReport();

            bool printResult = false;
            MethodInvoker mi = delegate
            {
                Export(viewer.LocalReport);
                pageIndex = 0;
                printResult = Print();
            };

            host.Invoke(mi);

            if (printResult)
            {
                return true;
            }
            else
            {
                viewer.PrintDialog();
                return true;
            }

            //            if(Print())
            //            {
            //return true;
            //            }
            //            else 
            //            {
            //                viewer.PrintDialog();
            //                return true;
            //            }

        }


        static private Stream CreateStream(string name, string fileNameExtension, Encoding encoding,
                                  string mimeType, bool willSeek)
        {
            Stream stream = new FileStream(name + "." + fileNameExtension, FileMode.Create);
            reportStreams.Add(stream);
            return stream;
        }

        static private void Export(LocalReport report)
        {
            string deviceInfo =
              "<DeviceInfo>" +
              "  <OutputFormat>EMF</OutputFormat>" +
                //"  <PageWidth>8.5in</PageWidth>" +
                //"  <PageHeight>11in</PageHeight>" +
                //"  <MarginTop>0.25in</MarginTop>" +
                //"  <MarginLeft>0.25in</MarginLeft>" +
                //"  <MarginRight>0.25in</MarginRight>" +
                //"  <MarginBottom>0.25in</MarginBottom>" +
                "  <MarginTop>0in</MarginTop>" +
              "  <MarginLeft>0in</MarginLeft>" +
              "  <MarginRight>0in</MarginRight>" +
              "  <MarginBottom>0in</MarginBottom>" +
              "</DeviceInfo>";
            Warning[] warnings;
            reportStreams = new List<Stream>();
            report.Render("Image", deviceInfo, CreateStream, out warnings);

            foreach (Stream stream in reportStreams)
                stream.Position = 0;
        }

        static private void PrintPage(object sender, PrintPageEventArgs ev)
        {
            Metafile pageImage = new Metafile(reportStreams[pageIndex]);
            ev.Graphics.DrawImage(pageImage, ev.PageBounds);

            pageIndex++;
            ev.HasMorePages = (pageIndex < reportStreams.Count);
        }

        static private bool Print()
        {
            try
            {
                string printerName = "OneNote";// "Microsoft Office Document Image Writer";

                if (!string.IsNullOrEmpty(defaultPrinter))
                {
                    printerName = defaultPrinter;
                }

                if (reportStreams == null || reportStreams.Count == 0)
                    return false;

                PrintDocument printDoc = new PrintDocument();
                printDoc.PrinterSettings.PrinterName = printerName;
                if (!printDoc.PrinterSettings.IsValid)
                {
                    string msg = String.Format("Can't find printer \"{0}\".", printerName);
                    Console.WriteLine(msg);
                    return false;
                }
                printDoc.PrintPage += new PrintPageEventHandler(PrintPage);
                printDoc.Print();
                return true;
            }
            catch (System.Exception ex)
            {
                Utility.ShowInformation("打印错误:" + ex.Message);
                return true;
            }


        }


        //第90%位数以上为“上等”，第10％位数以下为“下等”。
        public static string GetHeightLevel(int year, float height, bool isBoy)
        {
            List<float> std;
            var dic = isBoy ? dicHB : dicHG;
            if (dic.TryGetValue(year, out std))
            {
                int index =std.FindLastIndex(s => height > s);
                if (index > -1)
                {
                    if (index > (int)PercentLevel.P90)
                    {
                        return "上等";
                    }
                    else if (index < (int)PercentLevel.P10)
                    {
                        return "下等";
                    }
                }
                else 
                {
                    return "下等";
                }
            }
            return "标准";

        }
        //第90%位数以上为“上等”，第10％位数以下为“下等”。
        public static string GetWeightLevel(int year, float weight, bool isBoy)
        {
            List<float> std;
            var dic = isBoy ? dicWB : dicWG;
            if (dic.TryGetValue(year, out std))
            {
                int index = std.FindLastIndex(s => weight > s);
                if (index > -1)
                {
                    if (index > (int)PercentLevel.P90)
                    {
                        return "上等";
                    }
                    else if (index < (int)PercentLevel.P10)
                    {
                        return "下等";
                    }
                }
                else 
                {
                    return "下等";
                }
            }
            return "标准";
        }

        public static bool IsUnderNutritionLevel(int year, bool isBoy, float weight, float height)
        {
            Dictionary<float, List<float>> dicOfYear;
            var dic = isBoy ? dicNutBoyEx : dicNutGirlEx;

            if (dic.TryGetValue(year, out dicOfYear))
            {
                foreach (var kp in dicOfYear)
                {
                    float min, max;
                    min = kp.Value[0];
                    max = kp.Value[1];
                    if(height>=min && height <=max)
                    {
                        bool ret = weight < kp.Key;
            //            Console.WriteLine(" 年龄={0},性别={1},身高={2}cm,体重={3}kg,是否营养不良={4}",
               //             year, isBoy ? "男" : "女", height, weight, ret ? "是" : "否");
                        return ret;
                    }
                }
            }
            //Console.WriteLine("未找到匹配标准");
            return true;

        }

        public static bool IsPinXue(int year, bool isboy, float xhs)
        {
            if(year>=5&&year<=11)
            {
                return xhs < 115;
            }
            else if(year >=12&&year<=14)
            {
                return xhs < 120;
            }
            else 
            {
                if(isboy)
                {
                    return xhs < 130;
                }
                else 
                {
                    return xhs < 120;
                }
            }
            return false;
        }


        //受试学生体重在标准体重之90％～110％之间为营养状况良好，低于90％为营养不良（分为轻、中、重度），超过120％为肥胖。
        public static string GetNutritionLevel(int year, float height, float weight, bool isBoy)
        {
            //        static Dictionary<int, Dictionary<int, List<float>>> dicNutGirl = new Dictionary<int, Dictionary<int, List<float>>>();
            Dictionary<int, List<float>> dicOfYear;
            var dic = isBoy ? dicNutBoy : dicNutGirl;

            if(dic.TryGetValue(year,out dicOfYear))
            {
                int intH = (int)(height + .5);

                List<float> list;
                if(dicOfYear.TryGetValue(intH,out list))
                {
                    //营养不良 和 肥胖
               
                    if (weight > list[5]) return "肥胖";
                    if (weight > list[4]) return "超重";
                    if (weight < list[3]) return "极重";//营养不良
                    if (weight < list[2]) return "重";
                    if (weight < list[1]) return "中";
                    if (weight < list[0]) return "轻";
                }
            }
            return "标准";
        }

        public static void OpenFileLocations(Dictionary<string, string> locations)
        {
            if (locations == null || locations.Count == 0)
            {
                return;
            }

            foreach (string fileName in locations.Values)
            {
                // add the quotation mark to make sure it's working when name contains comma
                string args = string.Format("/select, \"{0}\"", fileName);
                Process p = new Process();
                p.StartInfo.FileName = "explorer.exe";
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.Arguments = args;
                p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.RedirectStandardInput = true;
                p.StartInfo.RedirectStandardOutput = true;
                p.StartInfo.RedirectStandardError = true;
                p.Start();

                p.WaitForExit(100);
                p.Close();
            }
        }

        private static void ReadStandard()
        {
            Assembly asm = Assembly.GetExecutingAssembly();

            string ps, ns, nns;
            ps = ns = nns = string.Empty;

            foreach (string str in asm.GetManifestResourceNames())
            {
                if (str.StartsWith("SchoolHealth.PERCENT"))
                {
                    ps = str; continue;
                }
                if (str.StartsWith("SchoolHealth.NUTRITION"))
                {
                    ns = str; continue;
                }
                if (str.StartsWith("SchoolHealth.NEWNUTRITION"))
                {
                    nns = str; continue;
                }
            }

            const string h = "HEIGHT";
            const string w = "WEIGHT";
            const string b = "BOY";
            const string g = "GIRL";

            bool hb, hg, wb, wg;
            hb = hg = wb = wg = false;

            using (TextReader pstr = new StreamReader(asm.GetManifestResourceStream(ps)))
            {
                var l = pstr.ReadLine();
                while (!string.IsNullOrEmpty(l))
                {
                    if (l.StartsWith(h) && l.EndsWith(b))
                    {
                        hb = true; hg = wb = wg = false;
                    }
                    else if (l.StartsWith(h) && l.EndsWith(g))
                    {
                        hg = true; hb = wb = wg = false; ;
                    }
                    else if (l.StartsWith(w) && l.EndsWith(b))
                    {
                        wb = true; hb = hg = wg = false;
                    }
                    else if (l.StartsWith(w) && l.EndsWith(g))
                    {
                        wg = true; hb = hg = wb = false;
                    }
                    else
                    {
                        var strs = l.Split(new string[] { " ", "\t" }, StringSplitOptions.RemoveEmptyEntries);
                        int y = int.Parse(strs[0]);
                        List<float> vals = new List<float>();
                        for (int i = 1; i < strs.Length; i++)
                        {
                            vals.Add(float.Parse(strs[i]));
                        }
                        Dictionary<int, List<float>> dic = null;

                        if (hb)
                        {
                            dic = dicHB;
                        }
                        else if (hg)
                        {
                            dic = dicHG;
                        }
                        else if (wb)
                        {
                            dic = dicWB;
                        }
                        else if (wg)
                        {
                            dic = dicWG;
                        }

                        if (dic != null)
                            dic[y] = vals;
                    }

                    l = pstr.ReadLine();
                }

            }

            int boyMin, boyMax, girlMin, girlMax;
            boyMin = boyMax = girlMin = girlMax = 0;


            using (TextReader pstr = new StreamReader(asm.GetManifestResourceStream(ns)))
            {
                var l = pstr.ReadLine();
                while (!string.IsNullOrEmpty(l))
                {
                    if (l.StartsWith(b))
                    {
                        var strs = l.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

                        if (strs.Length == 4)
                        {
                            boyMin = boyMax = int.Parse(strs[1]);
                            girlMin = girlMax = int.Parse(strs[3]);
                        }
                        else if (strs.Length == 6)
                        {
                            boyMin = int.Parse(strs[1]);
                            boyMax = int.Parse(strs[2]);
                            girlMin = int.Parse(strs[4]);
                            girlMax = int.Parse(strs[5]);
                        }
                    }
                    else
                    {
                        var strs = l.Split(new string[] { " ", "\t" }, StringSplitOptions.RemoveEmptyEntries);
                        int ht = int.Parse(strs[0]);
                        List<float> vals = new List<float>();
                        for (int i = 1; i < strs.Length; i++)
                        {
                            vals.Add(float.Parse(strs[i]));
                        }

                        for (int i = boyMin; i <= boyMax; i++)
                        {
                            Dictionary<int, List<float>> dic;

                            if (!dicNutBoy.TryGetValue(i, out dic))
                            {
                                dic = new Dictionary<int, List<float>>();
                                dicNutBoy[i] = dic;
                            }
                            dic[ht] = vals;
                        }

                        for (int i = girlMin; i <= girlMax; i++)
                        {
                            Dictionary<int, List<float>> dic;

                            if (!dicNutGirl.TryGetValue(i, out dic))
                            {
                                dic = new Dictionary<int, List<float>>();
                                dicNutGirl[i] = dic;
                            }
                            dic[ht] = vals;
                        }

                    }

                    l = pstr.ReadLine();
                }
            }

            //new standard
            bool isBoy;
            int minOld, maxOld;
            minOld = maxOld = 0;
            isBoy = true;
            


            using (TextReader pstr = new StreamReader(asm.GetManifestResourceStream(nns)))
            {
                var l = pstr.ReadLine();
                while (!string.IsNullOrEmpty(l))
                {
                    if (l.StartsWith(b) || l.StartsWith(g))
                    {
                        isBoy = l.StartsWith(b);
                        var strs = l.Split(new string[] { " ", "\t" }, StringSplitOptions.RemoveEmptyEntries);
                        if (strs.Length == 3)
                        {
                            minOld = int.Parse(strs[1]);
                            maxOld = int.Parse(strs[2]);
                        }
                    }
                    else
                    {
                        var strs = l.Split(new string[] { " ", "\t" }, StringSplitOptions.RemoveEmptyEntries);
                        float hmin, hmax, weight;

                        hmin = float.Parse(strs[0]);
                        hmax = float.Parse(strs[1]);
                        weight = float.Parse(strs[2]);

                        for (int i = minOld; i <= maxOld; i++)
                        {
                            Dictionary<float, List<float>> dic;
                            var tmp = isBoy ? dicNutBoyEx : dicNutGirlEx;

                            if (!tmp.TryGetValue(i, out dic))
                            {
                                dic = new Dictionary<float, List<float>>();
                                tmp[i] = dic;
                            }
                            dic[weight] = new List<float> { hmin, hmax };
                        }
                    }

                    l = pstr.ReadLine();
                }
            }




        }
    }
    internal enum ExportType
    {
        Excel = 0,
        Word = 3,
    }
    internal enum PercentLevel
    {
        P3,
        P5,
        P10,
        P15,
        P25,
        P30,
        P50,
        P70,
        P75,
        P85,
        P90,
        P95,
        P97
    }

    internal enum NutritionLevel
    {
        LittleThin, MiddleThin, Thin, VeryThin, Fat, VeryFat
    }
}
