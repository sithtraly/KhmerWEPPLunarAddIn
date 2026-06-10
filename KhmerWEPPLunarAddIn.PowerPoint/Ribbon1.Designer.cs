namespace KhmerWEPPLunarAddIn.PowerPoint
{
	partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public Ribbon1()
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
			this.tab1 = this.Factory.CreateRibbonTab();
			this.GroupHome = this.Factory.CreateRibbonGroup();
			this.button1 = this.Factory.CreateRibbonButton();
			this.tab1.SuspendLayout();
			this.GroupHome.SuspendLayout();
			this.SuspendLayout();
			// 
			// tab1
			// 
			this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.tab1.ControlId.OfficeId = "TabHome";
			this.tab1.Groups.Add(this.GroupHome);
			this.tab1.Name = "tab1";
			// 
			// GroupHome
			// 
			this.GroupHome.Items.Add(this.button1);
			this.GroupHome.Label = "ប្រតិទិន";
			this.GroupHome.Name = "GroupHome";
			// 
			// button1
			// 
			this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.button1.Image = global::KhmerWEPPLunarAddIn.PowerPoint.Properties.Resources.icon_80;
			this.button1.Label = "ចន្ទគតិ";
			this.button1.Name = "button1";
			this.button1.ShowImage = true;
			this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
			// 
			// Ribbon1
			// 
			this.Name = "Ribbon1";
			this.RibbonType = "Microsoft.PowerPoint.Presentation";
			this.Tabs.Add(this.tab1);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
			this.tab1.ResumeLayout(false);
			this.tab1.PerformLayout();
			this.GroupHome.ResumeLayout(false);
			this.GroupHome.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup GroupHome;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
	}

	partial class ThisRibbonCollection
	{
		internal Ribbon1 Ribbon1
		{
			get { return this.GetRibbon<Ribbon1>(); }
		}
	}
}
