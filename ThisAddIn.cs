using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools;
using Microsoft.Win32;

namespace KhmerWEPPLunarAddIn
{
	public partial class ThisAddIn
	{
		public CustomTaskPane taskpane;
		private WEPP_Sidepane userControl;
		private void ThisAddIn_Startup(object sender, System.EventArgs e)
		{
			SetBrowserEmulationToIE11();

			userControl = new WEPP_Sidepane();
			taskpane = this.CustomTaskPanes.Add(userControl, "ប្រតិទិនចន្ទគតិ");
			taskpane.Width = 352;
			taskpane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
			taskpane.Visible = false;
		}

		private void SetBrowserEmulationToIE11()
		{
			try
			{
				string registryKeyPath = @"Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION";
				using (RegistryKey key = Registry.CurrentUser.OpenSubKey(registryKeyPath, true))
				{
					key?.SetValue("WINWORD.EXE", 11001, RegistryValueKind.DWord);
				}
			} catch (Exception ex) {
				System.Diagnostics.Debug.WriteLine("Registry fix failed: " + ex.Message);
			}
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
