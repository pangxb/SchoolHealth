using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using DevComponents.DotNetBar.Validator;

namespace SchoolHealth
{
    public partial class SplashForm : Form
    {
        int rotator;
        public SplashForm()
        {
            InitializeComponent();
            rotator = 0;
            timer1.Enabled = true;
        }

        public void UpdateText(string msg)
        {
            MethodInvoker mi = delegate
            {
                labelX1.Text = msg;
            };
            if (this.IsHandleCreated && !this.IsDisposed)
            {
                this.Invoke(mi);
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            eHighlightColor color = (eHighlightColor)(rotator++ % 4);
           
            if(color == eHighlightColor.Red)
            {
                Random r = new Random((int)DateTime.Now.Ticks);
                int index = r.Next(2, 4);
                color = (eHighlightColor)index;
            }

           // highlighter1.SetHighlightColor(progressBarX1, color);
        }

        private void panelEx1_Click(object sender, EventArgs e)
        {

        }
    }
}