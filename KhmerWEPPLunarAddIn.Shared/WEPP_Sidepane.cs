using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace KhmerWEPPLunarAddIn.Shared
{
	[ComVisible(true)]
	public partial class WEPP_Sidepane : UserControl
	{
		private readonly IOfficeBridge _officeBridge;
		public WEPP_Sidepane(IOfficeBridge officeBridge)
		{
			InitializeComponent();
			_officeBridge = officeBridge;

			// Link the JavaScript 'window.external' object directly to this C# class
			webBrowser1.ObjectForScripting = this;
			InitializeWebView();
		}

		private async void InitializeWebView()
		{
			string resourceName = "KhmerWEPPLunarAddIn.Shared.Resources.index.html";
			Assembly assembly = Assembly.GetExecutingAssembly();
			using (Stream stream = assembly.GetManifestResourceStream(resourceName))
			{
				if (stream == null) throw new Exception($"Resource '{resourceName}' not found.");

				using (StreamReader reader = new StreamReader(stream))
				{
					string htmlContent = reader.ReadToEnd();
					webBrowser1.DocumentText = htmlContent;
				}
			}
		}

		public void converted(string msg)
		{
			//Microsoft.Office.Interop.Word.Selection selection = Globals.ThisAddIn.Application.Selection;
			//selection.TypeText(msg);
			_officeBridge.InsertText(msg);
		}
	}
}
