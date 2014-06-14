using System;
using System.Collections.Generic;

using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;
using System.Reflection;
using DevComponents.DotNetBar;
using System.Drawing;

namespace SchoolHealth
{
    static class Program
    {
       static AppSingletonChecker appSingletonChecker;
       
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            int i = 10000;

            Console.WriteLine("{0:d4}",i);
            appSingletonChecker = new AppSingletonChecker();
            if(!appSingletonChecker.IsSingle)
            {
                Utility.ShowInformation("已经运行一个程序");
                appSingletonChecker.Dispose();
                return;
            }
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
            Application.ThreadException += new ThreadExceptionEventHandler(Application_ThreadException);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                SplashAppContext context = SplashAppContext.GlobalSplashAppContext;
                Application.Run(context);
            }
            catch(Exception ex)
            {
                Utility.ShowInformation(ex.Message);
            }
            finally
            {
                appSingletonChecker.Dispose();
            }
        }

        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            
        }

        static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            
        }
    }


    public class AppSingletonChecker : IDisposable
    {
        #region Data Members

        private Mutex mutex;
        private bool isNew = false;

        private static readonly string AppGuid = "AE9F3C36-FA00-4178-AF8C-1E4FB3EAFF1B";

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor to get process name
        /// </summary>
        /// <param name="name">name of process/application</param>
        public AppSingletonChecker(string name)
        {
            string mutexName = Assembly.GetExecutingAssembly().GetName().Name + name;
            mutex = new Mutex(true, mutexName, out isNew);
        }

        public AppSingletonChecker()
        {
            mutex = new Mutex(true, AppGuid, out isNew);
        }

        #endregion

        #region Destructor

        /// <summary>
        /// Destructor to release mutex
        /// </summary>
        ~AppSingletonChecker()
        {
            ReleaseMutex();
        }

        #endregion

        #region Properties

        /// <summary>
        /// get true it is first instance
        /// </summary>
        public bool IsSingle
        {
            get { return isNew; }
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Relase owned mutex
        /// </summary>
        private void ReleaseMutex()
        {
            if (isNew)
            {
                mutex.ReleaseMutex();
                isNew = false;
            }
        }

        #endregion

        #region IDisposable Members

        /// <summary>
        /// dispose and release mutex
        /// </summary>
        public void Dispose()
        {
            ReleaseMutex();
            GC.SuppressFinalize(this);
        }

        #endregion
    }


}
