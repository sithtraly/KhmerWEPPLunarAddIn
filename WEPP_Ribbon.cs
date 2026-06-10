using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KhmerWEPPLunarAddIn
{
	public partial class WEPP_Ribbon
	{
		private void WEPP_Ribbon_Load(object sender, RibbonUIEventArgs e)
		{

		}

		private void BtOpenSidepane_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.taskpane.Width = 420;
			Globals.ThisAddIn.taskpane.Visible = true;
		}
	}
}
