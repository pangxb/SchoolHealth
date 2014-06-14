//===============================================================================
// Microsoft patterns & practices
// CompositeUI Application Block
//===============================================================================
// Copyright ?Microsoft Corporation.  All rights reserved.
// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY
// OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
// LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
// FITNESS FOR A PARTICULAR PURPOSE.
//===============================================================================

using System;
using System.ComponentModel;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using System.Drawing;

using System.Windows;
using System.Threading;
using System.Diagnostics;


namespace SchoolHealth
{
    public partial class ProgressForm : Office2007Form
    {
        private static readonly string progressTextFormat = "{0}%";
        private bool bReportProgress;
        private bool bShowProgressText;
        private bool bSupportCancel;
        private bool isAutoCloseAfterCompleted;
        private bool isOperationCompleted;
        private bool bHidePicture = false;

        private string msgTextFormat;
        private object argument;
        private IWin32Window owner;
        private BackgroundWorker worker = null;

        public event EventHandler Canceled;

        public ProgressForm()
        {
            InitializeComponent();
            isAutoCloseAfterCompleted = true;
            worker = new BackgroundWorker();
       //     lblMessage.ForeColor = Utility.GloablColorSchemeItem.ForeColor;

            try
            {
                if (!this.IsHandleCreated)
                {
                    this.CreateHandle();
                }
            }
            catch
            {
                this.RecreateHandle();
            }

            this.FormClosed += new FormClosedEventHandler(OnFormClosed);

            //if(Environment.OSVersion.Version.Major >=6)
            //{   
            //    this.FormBorderStyle = FormBorderStyle.FixedToolWindow;

            //}
        }

        public bool HidePicture
        {
            get { return bHidePicture; }
            set
            {
                bHidePicture = value;
                this.pbProcess.Visible = !bHidePicture;
                ShowHidePicutre(bHidePicture);
            }
        }

        private void ShowHidePicutre(bool bHidePicture)
        {
            //if (bHidePicture)
            //{
            //    if (this.panelEx1.Controls.Contains(this.pbProcess))
            //        this.panelEx1.Controls.Remove(this.pbProcess);

            //    this.pbProcess.Image = null;
            //}
            //else
            //{
            //    if (!this.panelEx1.Controls.Contains(this.pbProcess))
            //        this.panelEx1.Controls.Add(this.pbProcess);

            //    this.pbProcess.Image = global::EDM.UserControls.Properties.Resources.processing;
            //}
        }

        public IWin32Window OwnerForm
        {
            set
            {
                owner = value;
            }
        }

        public bool ShowProgressText
        {
            set
            {
                if (InvokeRequired && IsHandleCreated && !IsDisposed)
                {
                    Utility.SaveInvoke(this, delegate
                   {
                       bShowProgressText = value;
                       progressBar.TextVisible = true;
                       lblMessage.Visible = true;
                   }, false, false);
                }
                else
                {
                    bShowProgressText = value;
                    progressBar.TextVisible = true;
                    lblMessage.Visible = true;
                }
            }
        }

        public object Argument
        {
            get { return argument; }
            set { argument = value; }
        }

        public string MessagFormat
        {
            set { msgTextFormat = value; }
        }


        public bool IsShowTextOnly
        {
            set
            {
                if (InvokeRequired && IsHandleCreated && !IsDisposed)
                {
                    Utility.SaveInvoke(this, delegate
                        {
                            if (value && !bHidePicture)
                                this.pbProcess.Visible = true;
                            else
                                this.pbProcess.Visible = false;

                            ShowHidePicutre(!value || bHidePicture);

                            this.progressBar.Visible = !value;
                            if (value)
                            {
                                this.Height = (int)(this.Height * 2 / 3);
                                //this.lblMessage.TextLineAlignment = StringAlignment.Center;
                                //this.lblMessage.TextAlignment = StringAlignment.Center;
                                this.lblMessage.Location = new Point(40, this.Height / 4);
                                this.lblMessage.Size = new Size(lblMessage.Size.Width - 40, lblMessage.Height);
                            }
                        }, false, false);
                }
                else
                {
                    if (value && !bHidePicture)
                        this.pbProcess.Visible = true;
                    else
                        this.pbProcess.Visible = false;

                    ShowHidePicutre(!value || bHidePicture);

                    this.progressBar.Visible = !value;
                    if (value)
                    {
                        this.Height = (int)(this.Height * 2 / 3);
                        //this.lblMessage.TextLineAlignment = StringAlignment.Center;
                        //this.lblMessage.TextAlignment = StringAlignment.Center;
                        this.lblMessage.Location = new Point(40, this.Height / 4);
                        this.lblMessage.Size = new Size(lblMessage.Size.Width - 40, lblMessage.Height);
                    }
                }
                //this.lblMessage.TextAlign = ContentAlignment.MiddleLeft;
            }
        }

        public bool IsReportProgress
        {
            get
            {
                return bReportProgress;
            }
            set
            {
                if (InvokeRequired && IsHandleCreated && !IsDisposed)
                {
                    Utility.SaveInvoke(this, delegate
                    {
                        bReportProgress = value;
                        worker.WorkerReportsProgress = bReportProgress;
                        if (bReportProgress)
                        {
                            progressBar.ProgressType = eProgressItemType.Standard;
                            worker.ProgressChanged -= new ProgressChangedEventHandler(OnProgressChanged);
                            worker.ProgressChanged += new ProgressChangedEventHandler(OnProgressChanged);
                        }
                        else
                        {
                            progressBar.ProgressType = eProgressItemType.Marquee;
                        }
                    }, false, false);
                }
                else
                {
                    bReportProgress = value;
                    worker.WorkerReportsProgress = bReportProgress;
                    if (bReportProgress)
                    {
                        progressBar.ProgressType = eProgressItemType.Standard;
                        worker.ProgressChanged -= new ProgressChangedEventHandler(OnProgressChanged);
                        worker.ProgressChanged += new ProgressChangedEventHandler(OnProgressChanged);
                    }
                    else
                    {
                        progressBar.ProgressType = eProgressItemType.Marquee;
                    }
                }
            }
        }

        public bool CancellationPending
        {
            get
            {
                if (worker == null)
                {
                    return false;
                }
                return worker.CancellationPending;
            }
        }

        public bool IsBusy
        {
            get
            {
                if (worker == null)
                {
                    return false;
                }
                return worker.IsBusy;
            }
        }

        public bool IsAutoCloseAfterCompleted
        {
            get
            {
                return isAutoCloseAfterCompleted;
            }
            set
            {
                isAutoCloseAfterCompleted = value;
            }
        }

        public bool IsOperationCompleted
        {
            get
            {
                return isOperationCompleted;
            }
            set
            {
                isOperationCompleted = value;
            }
        }

        public void ReportProgress(int percent, object status)
        {
            try
            {
                if (InvokeRequired && IsHandleCreated && !IsDisposed)
                {
                    MethodInvoker mi = delegate
                    {
                        if (bReportProgress && worker.IsBusy && this.IsHandleCreated && !this.IsDisposed && !this.Disposing)
                        {
                            worker.ReportProgress(percent, status);
                        }
                    };
                    Utility.SaveInvoke(this, mi, false, false);
                }
                else
                {
                    if (bReportProgress && worker.IsBusy && this.IsHandleCreated && !this.IsDisposed && !this.Disposing)
                    {
                        worker.ReportProgress(percent, status);
                    }
                }
            }
            catch (Exception ex)
            {
          //      Utils.Utility.Log(ex);
            }
        }

        public bool SupportCancel
        {
            set
            {
                if (InvokeRequired && IsHandleCreated && !IsDisposed)
                {
                    MethodInvoker mi = delegate
                    {
                        bSupportCancel = value;
                        worker.WorkerSupportsCancellation = bSupportCancel;
                        btnCancel.Visible = bSupportCancel;
                    };
                    Utility.SaveInvoke(this, mi, false, false);
                }
                else
                {
                    bSupportCancel = value;
                    worker.WorkerSupportsCancellation = bSupportCancel;
                    btnCancel.Visible = bSupportCancel;
                }
            }
        }


        public event RunWorkerCompletedEventHandler RunWorkerCompleted
        {
            add { worker.RunWorkerCompleted += value; }
            remove { worker.RunWorkerCompleted -= value; }
        }

        public event DoWorkEventHandler DoWork
        {
            add { worker.DoWork += value; }
            remove { worker.DoWork -= value; }
        }

        public event ProgressChangedEventHandler ProgressChanged
        {
            add { worker.ProgressChanged += value; }
            remove { worker.ProgressChanged -= value; }
        }

        public void OnProcessCompleted(object sender, EventArgs e)
        {
            DoEventsTimer.Stop();

            bReportProgress = false;


            if (isAutoCloseAfterCompleted)
            {
                Close();
            }

            //if (this.owner != null)
            //{
            //    Form form = this.owner as Form;
            //    if (form != null && form.Visible)
            //    {
            //        form.BringToFront();
            //    }
            //}
        }

        public void OnProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            try
            {
                if (InvokeRequired && IsHandleCreated && !IsDisposed)
                {
                    MethodInvoker mi = delegate
                    {
                        progressBar.Text = string.Format(progressTextFormat, e.ProgressPercentage);

                        //if(e.ProgressPercentage > progressBar.Value)
                        //{ 
                        progressBar.Value = e.ProgressPercentage;
                        //}
                        if (e.UserState != null)
                        {
                            if (!string.IsNullOrEmpty(msgTextFormat))
                            {
                                lblMessage.Text = string.Format(msgTextFormat, e.UserState);
                            }
                            else
                            {
                                lblMessage.Text = e.UserState.ToString();
                            }
                            lblMessage.Refresh();
                            progressBar.Refresh();
                        }
                        //Application.DoEvents();
                    };
                    Utility.SaveInvoke(this, mi, false, false);
                }
                else
                {
                    progressBar.Text = string.Format(progressTextFormat, e.ProgressPercentage);

                    //if(e.ProgressPercentage > progressBar.Value)
                    //{ 
                    progressBar.Value = e.ProgressPercentage;
                    //}
                    if (e.UserState != null)
                    {
                        if (!string.IsNullOrEmpty(msgTextFormat))
                        {
                            lblMessage.Text = string.Format(msgTextFormat, e.UserState);
                        }
                        else
                        {
                            lblMessage.Text = e.UserState.ToString();
                        }
                        lblMessage.Refresh();
                        progressBar.Refresh();
                    }
                    //Application.DoEvents();
                }
            }
            catch
            {

            }

        }

        //public void Start()
        //{
        //    worker.RunWorkerCompleted -= new RunWorkerCompletedEventHandler(OnProcessCompleted);
        //    worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(OnProcessCompleted);
        //    worker.RunWorkerAsync();
        //    //this.BringToFront();
        //    if (owner != null)
        //    {
        //        this.StartPosition = FormStartPosition.CenterParent;
        //        ShowDialog(owner);
        //    }
        //    else
        //    {
        //        this.StartPosition = FormStartPosition.CenterScreen;
        //        ShowDialog();
        //    }
        //}

        public new string Text
        {
            get
            {
                return this.lblMessage.Text;
            }
            set
            {
                if (InvokeRequired && IsHandleCreated && !IsDisposed)
                {
                    MethodInvoker mi = delegate
                    {
                        this.lblMessage.Text = value;
                    };
                    Utility.SaveInvoke(this, mi, false, false);
                }
                else
                {
                    this.lblMessage.Text = value;
                }
            }
        }


        //public void Start(object arg)
        //{
        //    worker.RunWorkerCompleted -= new RunWorkerCompletedEventHandler(OnProcessCompleted);
        //    worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(OnProcessCompleted);
        //    if (arg != null)
        //    {
        //        argument = arg;
        //        worker.RunWorkerAsync(arg);
        //        if (owner != null)
        //        {
        //            this.StartPosition = FormStartPosition.CenterParent;
        //            ShowDialog(owner);
        //        }
        //        else
        //        {
        //            this.StartPosition = FormStartPosition.CenterScreen;
        //            ShowDialog();
        //        }
        //    }
        //    else
        //    {
        //        Start();
        //    }
        //}


        private System.Windows.Forms.Timer doEventsTimer;

        private System.Windows.Forms.Timer DoEventsTimer
        {
            get
            {
                if (doEventsTimer == null)
                {
                    doEventsTimer = new System.Windows.Forms.Timer();
                    doEventsTimer.Interval = 200;
                    doEventsTimer.Tick += (OnDoEventsTimerTick);
                }
                return doEventsTimer;
            }
        }

        void OnDoEventsTimerTick(object sender, EventArgs e)
        {
            Application.DoEvents();
        }

        public void Start(object arg = null, IWin32Window owner = null)
        {
            worker.RunWorkerCompleted -= (OnProcessCompleted);
            worker.RunWorkerCompleted += (OnProcessCompleted);

            if (owner != null)
            {
                this.owner = owner;
            }

            DoEventsTimer.Start();

            if (arg != null)
            {
                argument = arg;
                worker.RunWorkerAsync(arg);
            }
            else
            {
                worker.RunWorkerAsync();
            }

            if (this.owner != null)
            {
                Form form = this.owner as Form;

                if (form != null && form.WindowState != FormWindowState.Minimized)
                {
                    this.StartPosition = FormStartPosition.CenterParent;
                }
                else
                {
                    this.StartPosition = FormStartPosition.CenterScreen;
                }

                ShowDialog(this.owner);
            }
            else
            {
                this.StartPosition = FormStartPosition.CenterScreen;
                ShowDialog();
            }

        }

        public void Start(bool visible, object arg = null, IWin32Window owner = null)
        {
            worker.RunWorkerCompleted -= (OnProcessCompleted);
            worker.RunWorkerCompleted += (OnProcessCompleted);

            if (owner != null)
            {
                this.owner = owner;
            }

            DoEventsTimer.Start();

            if (arg != null)
            {
                argument = arg;
                worker.RunWorkerAsync(arg);
            }
            else
            {
                worker.RunWorkerAsync();
            }

            if (this.owner != null)
            {
                this.StartPosition = FormStartPosition.CenterParent;
                if (visible)
                    ShowDialog(this.owner);
            }
            else
            {
                this.StartPosition = FormStartPosition.CenterScreen;
                if (visible)
                    ShowDialog();
            }

        }

        public bool CancelButtonVisible
        {
            set
            {
                if (InvokeRequired && IsHandleCreated && !IsDisposed)
                {
                    MethodInvoker mi = delegate
                    {
                        btnCancel.Visible = value;
                        this.Invalidate();
                    };
                    Utility.SaveInvoke(this, mi, false, false);
                }
                else
                {
                    btnCancel.Visible = value;
                    this.Invalidate();
                }
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (worker.IsBusy)
            {
                Text = "È¡ÏûÖÐ...";
                worker.CancelAsync();
            }
            if (Canceled != null)
            {
                Canceled(this, null);
            }
            this.DialogResult = DialogResult.Cancel;
            Close();
        }

        private void OnFormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                worker.RunWorkerCompleted -= OnProcessCompleted;
                worker.ProgressChanged -= OnProgressChanged;
                DoEventsTimer.Tick -= OnDoEventsTimerTick;
                FormClosed -= OnFormClosed;

                if (InvokeRequired && IsHandleCreated && !IsDisposed)
                {
                    Utility.SaveInvoke(this, delegate
                    {
                        var form = this.owner as Form;
                        if (form != null)
                        {
                            form.BringToFront();
                        }

                        this.Dispose();
                    }, false, false);
                }
                else
                {
                    var form = this.owner as Form;
                    if (form != null)
                    {
                        form.BringToFront();
                    }

                    this.Dispose();
                }
            }
            catch (Exception ex)
            {
          
            }
            finally
            {

            }
        }

        /// <summary>
        /// ShowWithoutActivation
        /// </summary>
        //protected override bool ShowWithoutActivation
        //{
        //    get { return true; }
        //}

        //protected override void WndProc(ref Message m)
        //{
        //    //if(Visible)
        //    if (Environment.OSVersion.Version.Major < Common.CommonConstant.VistaVersion)
        //    {
        //        if (m.Msg == (int)WindowsMessages.WM_MOUSEACTIVATE)
        //        {
        //            m.Result = new IntPtr((int)WindowsMessages.MA_NOACTIVATE);
        //            return;
        //        }
        //        else if (m.Msg == (int)WindowsMessages.WM_NCACTIVATE)
        //        {
        //            //yunliang:please do not use following codes,it will cause exception in aync operations .
        //            //if (((int)m.WParam & 0xFFFF) != (int)WindowsMessages.WA_INACTIVE)
        //            //{
        //            //    if (m.LParam != IntPtr.Zero)
        //            //    {
        //            //        User32.SetActiveWindow(m.LParam);
        //            //    }
        //            //    else
        //            //    {
        //            //        User32.SetActiveWindow(IntPtr.Zero);
        //            //    }
        //            //}
        //        }
        //    }
        //    base.WndProc(ref m);
        //}

        //need more overload method
        public static void QueueWorkingItem(MethodInvoker invoker, string msg, IWin32Window parent)
        {
            using (ProgressForm pf = new ProgressForm())
            {
                pf.IsShowTextOnly = true;
                pf.Text = msg;

                Form form = parent as Form;

                if (form != null)
                {
                    pf.OwnerForm = form;
                }
                else
                {
                    Control ctrl = parent as Control;
                    if (ctrl != null)
                    {
                        pf.OwnerForm = ctrl.TopLevelControl ?? ctrl;
                    }
                    else
                    {
                        pf.OwnerForm = ctrl;
                    }
                }


                AutoResetEvent evt = new AutoResetEvent(false);

                ThreadPool.QueueUserWorkItem(s =>
                {
                    while (!evt.WaitOne(10)) Application.DoEvents();
                });

                pf.DoWork += delegate
                {
                    try
                    {
                        Control ctrl = parent as Control;
                        if (ctrl != null)
                        {
                            ctrl.Invoke(invoker);
                        }
                        else
                        {
                            invoker.Invoke();
                        }
                    }
                    catch 
                    {
                        
                    }
                    finally
                    {
                        evt.Set();
                    }
                
                };
                pf.Start();
            }
        }



        public static void QueueWorkingItemEx(MethodInvoker invoker, string msg, IWin32Window parent)
        {
            using (ProgressForm pf = new ProgressForm())
            {
                pf.IsShowTextOnly = true;
                pf.Text = msg;
                pf.SupportCancel = true;

                Form form = parent as Form;

                if (form != null)
                {
                    pf.OwnerForm = form;
                }
                else
                {
                    Control ctrl = parent as Control;
                    if (ctrl != null)
                    {
                        pf.OwnerForm = ctrl.TopLevelControl ?? ctrl;
                    }
                    else
                    {
                        pf.OwnerForm = ctrl;
                    }
                }


        
                pf.DoWork += delegate
                {
                    invoker.Invoke();
                };
                pf.Start();
            }
        }
    }



    #region ProgressImageBox

    public class ProgressImageBox : UserControl
    {

        private Image Image { get; set; }
       
        public ProgressImageBox()
        {
            try
            {
                Image = Properties.Resources.processing;

                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.SupportsTransparentBackColor, true);

                if (!this.IsHandleCreated)
                {
                    this.CreateHandle();
                }
            }
            catch
            {
                this.RecreateHandle();
            }
        }

        protected override void OnVisibleChanged(EventArgs e)
        {
            if (Image != null)
            {
                if (this.Visible)
                {
                    if (ImageAnimator.CanAnimate(Image))
                    {
                        ImageAnimator.StopAnimate(Image, OnAnimate);
                        ImageAnimator.Animate(Image, OnAnimate);
                    }
                }
                else
                {
                    ImageAnimator.StopAnimate(Image, OnAnimate);
                }
            }

            base.OnVisibleChanged(e);
        }

        protected override void OnPaint(PaintEventArgs pe)
        {
            if (Image != null)
            {
                ImageAnimator.UpdateFrames();
                pe.Graphics.DrawImage(this.Image, new Rectangle(0, 0, this.Width, this.Height));
            }

            base.OnPaint(pe);
        }

        protected override void Dispose(bool disposing)
        {
            try
            {
                if (Image != null)
                {
                    ImageAnimator.StopAnimate(Image, OnAnimate);

                    Image.Dispose();
                    Image = null;
                }

                if (IsHandleCreated)
                    DestroyHandle();
            }
            catch (Exception ex)
            {

                return;
            }
            finally
            {
                base.Dispose(disposing);
            }
        }

        private void OnAnimate(object sender, EventArgs e)
        {
            Utility.SaveInvoke(this, Invalidate, true, false);
        }

    }
    #endregion
}