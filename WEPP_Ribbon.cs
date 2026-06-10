using Microsoft.Office.Tools.Ribbon;

namespace KhmerWEPPLunarAddIn
{
	public partial class WEPP_Ribbon
	{
		private void WEPP_Ribbon_Load(object sender, RibbonUIEventArgs e)
		{

		}

		private void BtOpenSidepane_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.taskpane.Width = 410;
			Globals.ThisAddIn.taskpane.Visible = true;
		}
	}
}
