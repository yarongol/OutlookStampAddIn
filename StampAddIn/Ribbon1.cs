using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Globalization;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace StampAddIn
{
    [ComVisible(true)]
    public class StampRibbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public StampRibbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("StampAddIn.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, select the Ribbon XML item in Solution Explorer and then press F1

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnButtonClick(Office.IRibbonControl control)
        {
            Microsoft.Office.Interop.Outlook.Inspector
                myInspector = Globals.ThisAddIn.Application.ActiveInspector();

            // Only works in Word
            if (myInspector.EditorType == Outlook.OlEditorType.olEditorWord)
            {
                Word.Document wd = (Word.Document)myInspector.WordEditor;

                object start = wd.Application.Selection.Range.Start;
                object end = wd.Application.Selection.Range.End;
                Word.Range rng = wd.Range(ref start, ref end);

                // Get pre designated color.
                int iColor = Convert.ToInt32(Globals.ThisAddIn.strColor);
                rng.Text = " [" + Globals.ThisAddIn.strName + " - " + System.DateTime.Now.ToString("MMM d, yyyy", CultureInfo.CreateSpecificCulture("en-US")) + "] ";

                rng.Select();
                rng.Font.Color = (Microsoft.Office.Interop.Word.WdColor)iColor;
                rng.Start = rng.End;
                rng.Select();
            }
            else
            {
                MessageBox.Show("The Signiture add-in will only work when Outlook is using the Word editor. You can change your editor under Outlook options");
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
