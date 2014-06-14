using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SchoolHealth
{
    public partial class SHComboBox : DevComponents.DotNetBar.Controls.ComboBoxEx
    {
        protected override void OnMouseWheel(MouseEventArgs e)
        {
            if(this.Parent != null)
            {
                this.Parent.Focus();
            }
            System.Windows.Forms.HandledMouseEventArgs hme = (System.Windows.Forms.HandledMouseEventArgs)e;
            hme.Handled = true; 
        }
    }
}
