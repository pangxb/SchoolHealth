using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;

namespace SchoolHealth
{
    public partial class EDMLineControl : UserControl
    {
        public EDMLineControl()
        {
            InitializeComponent();
            Title = "Title";
            TitleFont = new Font(lblTitle.Font, FontStyle.Bold);
            base.MinimumSize = new Size(16, 17);
            base.MaximumSize = new Size(2048, 17);
        }

        [Browsable(true),Localizable(true)]
        public String Title 
        {
            get 
            {              
                return this.lblTitle.Text.Trim(); 
            }
            set
            {    
                if(string.IsNullOrEmpty(value))
                {
                    this.lblTitle.Text = string.Empty;
                    this.lblTitle.AutoSize = false;
                    if(lblTitle.Size.Width != 1)
                    {
                        var s = lblTitle.Size;
                        s.Width = 1;
                        this.lblTitle.Size = s;
                        base.MinimumSize = new Size(16, 8);
                        base.MaximumSize = new Size(2048, 8);
                    }
                }
                else 
                {
                    this.lblTitle.Text = string.Format("{0} ", value);
                    if(!lblTitle.AutoSize)
                    {
                        this.lblTitle.AutoSize = true;
                        base.MinimumSize = new Size(16, 17);
                        base.MaximumSize = new Size(2048, 17);
                    }

                }
                this.lblTitle.RecalcLayout();
            }
        }

          [Browsable(true)]
        public Font TitleFont
        {
            get { return lblTitle.Font; }
            set { lblTitle.Font = value; }
        }
        

        [Browsable(false)]
        public new Size MinimumSize
        {
            get { return base.MinimumSize; }
            set { return; }
        }

        [Browsable(false)]
        public new Size MaximumSize
        {
            get { return base.MaximumSize; }
            set { return; }
        }
    }
}
