
using Microsoft.Office.Tools;

namespace KhmerWEPPLunarAddIn.PowerPoint
{
	public partial class ThisAddIn
	{
		public CustomTaskPane taskpane;
		private void ThisAddIn_Startup(object sender, System.EventArgs e)
		{
			var bridge = new PowerPointBridge();
			var userControl = new KhmerWEPPLunarAddIn.Shared.WEPP_Sidepane(bridge);
			var taskPane = this.CustomTaskPanes.Add(userControl, "ប្រតិទិនចន្ទគតិ");
			taskPane.Visible = false;
			this.taskpane = taskPane;
		}

		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
		{
		}

		#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
			this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}

		#endregion
	}
}
