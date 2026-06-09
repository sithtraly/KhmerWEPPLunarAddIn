namespace KhmerWEPPLunarAddIn.Shared
{
	// This defines what JavaScript is allowed to request the host app to do
	public interface IOfficeBridge
	{
		void InsertText(string text);
	}
}