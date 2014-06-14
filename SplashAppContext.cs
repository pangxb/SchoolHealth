using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Threading;


namespace SchoolHealth
{
    public class SplashAppContext : ApplicationContext
    {
        private bool isFadeClose;
        private int showSplashInterval;
        private StartUpHelper helper;
        private Form mainForm;
        private System.Windows.Forms.Timer splashTimer;
        private System.Windows.Forms.Timer closeTimer;

       public static SplashAppContext GlobalSplashAppContext
       {
           get 
           {
               SplashForm sf = new SplashForm();
               SplashAppContext sac = new SplashAppContext(sf);
               return sac;
           }
       }

        public SplashAppContext(Form splashForm)
            : base(splashForm)
        {
            splashForm.Shown += (this.OnSplashFormShown);

            this.helper = new StartUpHelper();
            this.helper.SetStartUpCallbackDelegate((splashForm as SplashForm).UpdateText);

            this.splashTimer = new System.Windows.Forms.Timer();
            this.splashTimer.Tick += (OnSplashTimeUp);
            this.splashTimer.Interval = 100;
            this.splashTimer.Enabled = true;

            this.closeTimer = new System.Windows.Forms.Timer();
            this.closeTimer.Tick += (OnCloseTimerClick);
            this.closeTimer.Interval = 10;
            this.closeTimer.Enabled = true;
        }

        private void OnCloseTimerClick(object sender, EventArgs e)
        {
            if (helper.Wait(1))
            {
                SplashForm splashForm = base.MainForm as SplashForm;
                if (splashForm != null)
                {
                    while (splashForm.Opacity > 0)
                    {
                        splashForm.Opacity -= 0.1;
                        Thread.Sleep(10);
                    }
                    splashForm.Close();
                }
                this.closeTimer.Enabled = false;
            }
        }


        protected override void OnMainFormClosed(object sender, EventArgs e)
        {
            if (sender is SplashForm)
            {
                base.MainForm = this.mainForm;
                base.MainForm.Show();
            }
            else if (sender is MainForm)
            {
                base.OnMainFormClosed(sender, e);
            }
        }

        private void OnSplashFormShown(object sender, EventArgs e)
        {
            
            this.mainForm = new MainForm();
            if (mainForm != null)
            {
                //pangxb:todo...
                (mainForm as MainForm).Initialize(this.helper);
            }
        }

        private void OnSplashTimeUp(object sender, EventArgs e)
        {
            if (base.MainForm.Opacity < 1.0)
            {
                Form mainForm = base.MainForm;
                mainForm.Opacity += 0.05;
            }
        }

        public bool IsFadeOpen
        {
            set
            {
                if (value)
                {
                    base.MainForm.Opacity = 0.0;
                }
            }
        }
    }

    public class StartUpHelper
    {
        private AutoResetEvent resetEvent;
        private StartUpEventHandler startUpHandler;

        public StartUpHelper()
        {
            resetEvent = new AutoResetEvent(false);
        }

        internal void SetCompleted()
        {
            if (resetEvent != null)
            {
                resetEvent.Set();
            }
        }

        internal void SetStartUpCallbackDelegate(StartUpEventHandler handler)
        {
            startUpHandler = handler;
        }

        internal void UpdateMsg(string msg)
        {
            if (startUpHandler != null)
            {
                startUpHandler(msg);
            }
        }

        internal bool Wait(int timeout)
        {
            if (resetEvent != null)
            {
                return resetEvent.WaitOne(timeout, false);
            }
            return true;
        }
    }

    public delegate void StartUpEventHandler(string msg);
}
