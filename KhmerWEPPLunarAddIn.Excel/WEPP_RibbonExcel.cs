using Microsoft.Office.Tools.Ribbon;

namespace KhmerWEPPLunarAddIn.Excel
{
	public partial class WEPP_RibbonExcel
	{
		private void WEPP_RibbonExcel_Load(object sender, RibbonUIEventArgs e)
		{

		}

		private void button1_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.taskpane.Width = 410;
			Globals.ThisAddIn.taskpane.Visible = true;
		}
	}
}
