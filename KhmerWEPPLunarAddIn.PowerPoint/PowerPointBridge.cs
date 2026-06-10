using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using KhmerWEPPLunarAddIn.Shared;

namespace KhmerWEPPLunarAddIn.PowerPoint
{
	internal class PowerPointBridge : IOfficeBridge
	{
		public void InsertText(string text)
		{
			try
			{
				Application pptApp = Globals.ThisAddIn.Application;
				Presentation activePre = pptApp.ActivePresentation;
				Slide activeSlide = (Slide)pptApp.ActiveWindow.View.Slide;

				if (activeSlide != null)
				{
					Microsoft.Office.Interop.PowerPoint.Shape textBox = activeSlide.Shapes.AddTextbox(
						MsoTextOrientation.msoTextOrientationHorizontal,
						100, 100, 400, 100);
					textBox.TextFrame.TextRange.Text = text;
				}
			}
			catch (Exception e)
			{
				System.Diagnostics.Debug.WriteLine("PowerPoint Insertion Error: ");
				System.Diagnostics.Debug.WriteLine(e.Message);
			}
		}
	}
}
