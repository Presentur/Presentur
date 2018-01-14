using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using SharedPowerpointFavoritesPlugin.util;

namespace SharedPowerpointFavoritesPlugin
{
    [ComVisible(true)]
    public class MainRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private static readonly DebugLogger logger = DebugLogger.GetLogger(typeof(MainRibbon).Name);
        private ShapePersistence shapePersistence = ShapePersistence.INSTANCE;
        private ImportExportService importExportService = ImportExportService.INSTANCE;

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
            logger.Log("Open Button pressed.");
            SharedFavView.ShowOrFocus();
        }

        public void OnImportSharedFavButton(Office.IRibbonControl control)
        {
            logger.Log("Import button clicked.");
            if (!this.AskForImportConfirmation())
            {
                logger.Log("User cancelled import.");
                return;
            }
            var filePath = DialogUtil.GetFilePathViaDialog(isSaveAction: false);
            if (filePath != null)
            {
                if (this.importExportService.ImportFromFile(filePath))
                {
                    MessageBox.Show("Successfully imported favorites.");
                }
                else
                {
                    MessageBox.Show("An error occured while importing favorites.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public void OnExportSharedFavButton(Office.IRibbonControl control)
        {
            var filePath = DialogUtil.GetFilePathViaDialog(isSaveAction: true);
            if (filePath != null)
            {
                if (this.importExportService.ExportToFile(filePath))
                {
                    MessageBox.Show("Successfully exported favorites.");
                }
                else
                {
                    MessageBox.Show("An error occured while exporting favorites.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public void OnCopyFromClipboardButton(Office.IRibbonControl control)
        {
            if (!this.shapePersistence.SaveShapeFromClipBoard())
            {
                MessageBox.Show("Clipboard content could not be read.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                SharedFavView.ShowOrFocus();
            }
        }

        public void SaveFavoriteShape(Office.IRibbonControl control)
        {
            logger.Log("Save As FavoriteShape clicked.");
            Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            Shape selectedShape = selection.ShapeRange[1];
            if (selectedShape != null)
            {
                ShapePersistence.INSTANCE.SaveShape(selectedShape);
                SharedFavView.ShowOrFocus();
            }
            else
            {
                logger.Log("Could not save selection " + selection);
            }
        }

        private bool AskForImportConfirmation()
        {
            return DialogUtil.AskForConfirmation("Are you sure you want to import a favorites archive? This deletes your own favorites!");
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
