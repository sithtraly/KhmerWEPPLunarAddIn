using KhmerWEPPLunarAddIn.Shared;

namespace KhmerWEPPLunarAddIn
{
	public class WordBridge : IOfficeBridge
	{
		public void InsertText(string text)
		{
			// Word specific implementation
			var selection = Globals.ThisAddIn.Application.Selection;
			selection.TypeText(text);
		}
	}
}