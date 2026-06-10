using Microsoft.Office.Tools.Ribbon;

namespace KhmerWEPPLunarAddIn.PowerPoint
{
	public partial class Ribbon1
	{
		private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
		{

		}

		private void button1_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.taskpane.Width = 420;
			Globals.ThisAddIn.taskpane.Visible = true;
		}
	}
}
