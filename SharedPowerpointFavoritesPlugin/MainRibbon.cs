﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

// TODO:  Führen Sie diese Schritte aus, um das Element auf dem Menüband (XML) zu aktivieren:

// 1: Kopieren Sie folgenden Codeblock in die ThisAddin-, ThisWorkbook- oder ThisDocument-Klasse.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new MainRibbon();
//  }

// 2. Erstellen Sie Rückrufmethoden im Abschnitt "Menübandrückrufe" dieser Klasse, um Benutzeraktionen
//    zu behandeln, z.B. das Klicken auf eine Schaltfläche. Hinweis: Wenn Sie dieses Menüband aus dem Menüband-Designer exportiert haben,
//    verschieben Sie den Code aus den Ereignishandlern in die Rückrufmethoden, und ändern Sie den Code für die Verwendung mit dem
//    Programmmodell für die Menübanderweiterung (RibbonX).

// 3. Weisen Sie den Steuerelementtags in der Menüband-XML-Datei Attribute zu, um die entsprechenden Rückrufmethoden im Code anzugeben.  

// Weitere Informationen erhalten Sie in der Menüband-XML-Dokumentation in der Hilfe zu Visual Studio-Tools für Office.


namespace SharedPowerpointFavoritesPlugin
{
    [ComVisible(true)]
    public class MainRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public MainRibbon()
        {
        }

        #region IRibbonExtensibility-Member

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("SharedPowerpointFavoritesPlugin.MainRibbon.xml");
        }

        #endregion

        #region Menübandrückrufe
        //Erstellen Sie hier Rückrufmethoden. Weitere Informationen zum Hinzufügen von Rückrufmethoden finden Sie unter "http://go.microsoft.com/fwlink/?LinkID=271226".

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnOpenSharedFavButton(Office.IRibbonControl control)
        {
            Console.WriteLine("Open Button pressed.");
            SharedFavView.ShowOrFocus();
        }

        #endregion

        #region Hilfsprogramme

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
