namespace SchoolHealth
{
    partial class EDMLineControl
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.panelLine = new DevComponents.DotNetBar.PanelEx();
            this.panel = new System.Windows.Forms.Panel();
            this.lblTitle = new DevComponents.DotNetBar.LabelX();
            this.panel.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelLine
            // 
            this.panelLine.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.panelLine.CanvasColor = System.Drawing.SystemColors.Control;
            this.panelLine.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.panelLine.Location = new System.Drawing.Point(0, 11);
            this.panelLine.Name = "panelLine";
            this.panelLine.Size = new System.Drawing.Size(197, 2);
            this.panelLine.Style.Alignment = System.Drawing.StringAlignment.Center;
            this.panelLine.Style.BackColor1.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.panelLine.Style.BackColor2.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.panelLine.Style.Border = DevComponents.DotNetBar.eBorderType.SingleLine;
            this.panelLine.Style.BorderColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.panelLine.Style.BorderSide = DevComponents.DotNetBar.eBorderSide.Top;
            this.panelLine.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.panelLine.Style.GradientAngle = 90;
            this.panelLine.TabIndex = 1;
            // 
            // panel
            // 
            this.panel.Controls.Add(this.panelLine);
            this.panel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel.Location = new System.Drawing.Point(16, 0);
            this.panel.Name = "panel";
            this.panel.Size = new System.Drawing.Size(199, 25);
            this.panel.TabIndex = 2;
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            // 
            // 
            // 
            this.lblTitle.BackgroundStyle.Class = "";
            this.lblTitle.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblTitle.Dock = System.Windows.Forms.DockStyle.Left;
            this.lblTitle.Location = new System.Drawing.Point(0, 0);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.PaddingRight = 4;
            this.lblTitle.Size = new System.Drawing.Size(16, 17);
            this.lblTitle.TabIndex = 3;
            this.lblTitle.Text = "1";
            this.lblTitle.TextAlignment = System.Drawing.StringAlignment.Center;
            // 
            // EDMLineControl
            // 
            this.Controls.Add(this.panel);
            this.Controls.Add(this.lblTitle);
            this.Font = new System.Drawing.Font("Tahoma", 8.25F);
            this.Name = "EDMLineControl";
            this.Size = new System.Drawing.Size(215, 25);
            this.panel.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.PanelEx panelLine;
        private System.Windows.Forms.Panel panel;
        private DevComponents.DotNetBar.LabelX lblTitle;
    }
}
