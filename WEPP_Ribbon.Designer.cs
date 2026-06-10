namespace KhmerWEPPLunarAddIn
{
	partial class WEPP_Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public WEPP_Ribbon()
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
			this.weppTab = this.Factory.CreateRibbonTab();
			this.group1 = this.Factory.CreateRibbonGroup();
			this.btOpenSidepane = this.Factory.CreateRibbonButton();
			this.weppTab.SuspendLayout();
			this.group1.SuspendLayout();
			this.SuspendLayout();
			// 
			// weppTab
			// 
			this.weppTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.weppTab.ControlId.OfficeId = "TabHome";
			this.weppTab.Groups.Add(this.group1);
			this.weppTab.Name = "weppTab";
			// 
			// group1
			// 
			this.group1.Items.Add(this.btOpenSidepane);
			this.group1.Label = "ប្រតិទិន";
			this.group1.Name = "group1";
			this.group1.Position = this.Factory.RibbonPosition.AfterOfficeId("GroupEditing");
			// 
			// btOpenSidepane
			// 
			this.btOpenSidepane.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.btOpenSidepane.Image = global::KhmerWEPPLunarAddIn.Properties.Resources.icon_80;
			this.btOpenSidepane.Label = "ចន្ទគតិ";
			this.btOpenSidepane.Name = "btOpenSidepane";
			this.btOpenSidepane.ShowImage = true;
			this.btOpenSidepane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtOpenSidepane_Click);
			// 
			// WEPP_Ribbon
			// 
			this.Name = "WEPP_Ribbon";
			this.RibbonType = "Microsoft.Word.Document";
			this.Tabs.Add(this.weppTab);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.WEPP_Ribbon_Load);
			this.weppTab.ResumeLayout(false);
			this.weppTab.PerformLayout();
			this.group1.ResumeLayout(false);
			this.group1.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab weppTab;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton btOpenSidepane;
	}

	partial class ThisRibbonCollection
	{
		internal WEPP_Ribbon WEPP_Ribbon
		{
			get { return this.GetRibbon<WEPP_Ribbon>(); }
		}
	}
}
