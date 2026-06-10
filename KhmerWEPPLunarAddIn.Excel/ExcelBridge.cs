using KhmerWEPPLunarAddIn.Shared;
using Microsoft.Office.Interop.Excel;

namespace KhmerWEPPLunarAddIn.Excel
{
	public class ExcelBridge : IOfficeBridge
	{
		public void InsertText(string text)
		{
			string cell1 = text.Split('\n')[0];
			string cell2 = text.Split('\n')[1];
			// Excel specific implementation
			Window activeWindow = Globals.ThisAddIn.Application.ActiveWindow;
			Range activeCell = activeWindow.ActiveCell;
			if (activeCell != null)
			{
				activeCell.Value2 = cell1;
				activeCell.Offset[1, 0].Value2 = cell2;
			}
		}
	}
}