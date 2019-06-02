namespace Reply_And_Keep_Attachments
{
    partial class ReplyAndKeepAttachmentsRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ReplyAndKeepAttachmentsRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            this.replyAndKeepAttachmentsTab1 = this.Factory.CreateRibbonTab();
            this.replyAndKeepAttachmentsGroup1 = this.Factory.CreateRibbonGroup();
            this.replyAndKeepAttachmentsButton = this.Factory.CreateRibbonButton();
            this.replyAllAndKeepAttachmentsButton = this.Factory.CreateRibbonButton();
            this.replyAndKeepAttachmentsTab1.SuspendLayout();
            this.replyAndKeepAttachmentsGroup1.SuspendLayout();
            this.SuspendLayout();
            // 
            // replyAndKeepAttachmentsTab1
            // 
            this.replyAndKeepAttachmentsTab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.replyAndKeepAttachmentsTab1.Groups.Add(this.replyAndKeepAttachmentsGroup1);
            this.replyAndKeepAttachmentsTab1.Label = "Reply and Keep Attachments";
            this.replyAndKeepAttachmentsTab1.Name = "replyAndKeepAttachmentsTab1";
            // 
            // replyAndKeepAttachmentsGroup1
            // 
            this.replyAndKeepAttachmentsGroup1.Items.Add(this.replyAndKeepAttachmentsButton);
            this.replyAndKeepAttachmentsGroup1.Items.Add(this.replyAllAndKeepAttachmentsButton);
            this.replyAndKeepAttachmentsGroup1.Label = "Reply and Keep Attachments";
            this.replyAndKeepAttachmentsGroup1.Name = "replyAndKeepAttachmentsGroup1";
            // 
            // replyAndKeepAttachmentsButton
            // 
            this.replyAndKeepAttachmentsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.replyAndKeepAttachmentsButton.Image = global::Reply_And_Keep_Attachments.Properties.Resources.Reply_And_Keep_Attachments_Icon;
            this.replyAndKeepAttachmentsButton.Label = "Reply and Keep Attachments";
            this.replyAndKeepAttachmentsButton.Name = "replyAndKeepAttachmentsButton";
            this.replyAndKeepAttachmentsButton.ShowImage = true;
            this.replyAndKeepAttachmentsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ReplyAndKeepAttachmentsButton_Click);
            // 
            // replyAllAndKeepAttachmentsButton
            // 
            this.replyAllAndKeepAttachmentsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.replyAllAndKeepAttachmentsButton.Image = global::Reply_And_Keep_Attachments.Properties.Resources.Reply_All_And_Keep_Attachments_Icon;
            this.replyAllAndKeepAttachmentsButton.Label = "Reply All and Keep Attachments";
            this.replyAllAndKeepAttachmentsButton.Name = "replyAllAndKeepAttachmentsButton";
            this.replyAllAndKeepAttachmentsButton.ShowImage = true;
            this.replyAllAndKeepAttachmentsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ReplyAllAndKeepAttachmentsButton_Click);
            // 
            // ReplyAndKeepAttachmentsRibbon
            // 
            this.Name = "ReplyAndKeepAttachmentsRibbon";
            this.RibbonType = "Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.replyAndKeepAttachmentsTab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ReplyAndKeepAttachmentsRibbon_Load);
            this.replyAndKeepAttachmentsTab1.ResumeLayout(false);
            this.replyAndKeepAttachmentsTab1.PerformLayout();
            this.replyAndKeepAttachmentsGroup1.ResumeLayout(false);
            this.replyAndKeepAttachmentsGroup1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab replyAndKeepAttachmentsTab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup replyAndKeepAttachmentsGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton replyAndKeepAttachmentsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton replyAllAndKeepAttachmentsButton;
    }

    partial class ThisRibbonCollection
    {
        internal ReplyAndKeepAttachmentsRibbon ReplyAndKeepAttachmentsRibbon
        {
            get { return this.GetRibbon<ReplyAndKeepAttachmentsRibbon>(); }
        }
    }
}
