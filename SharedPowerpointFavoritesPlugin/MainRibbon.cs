/*
Presentur - Creating presentations in corporate design easily using your set of designed elements, icons and shapes.
Copyright (C) 2018 Christopher Rudoll, Eduard Hajek

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

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
using System.Drawing;
using SharedPowerpointFavoritesPlugin.model;

namespace SharedPowerpointFavoritesPlugin
{
    [ComVisible(true)]
    public class MainRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private static readonly DebugLogger logger = DebugLogger.GetLogger(typeof(MainRibbon).Name);
        private ShapePersistence shapePersistence = ShapePersistence.INSTANCE;
        private ImportExportService importExportService = ImportExportService.INSTANCE;
        private ShapeService shapeService = ShapeService.INSTANCE;
        private readonly int ITEM_WIDTH = 100;
        private readonly int ITEM_HEIGHT = 80;

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
            this.shapePersistence.RegisterCacheListener(new RibbonCacheListener(this));
        }

        class RibbonCacheListener : ShapePersistence.CacheListener
        {
            private MainRibbon mainRibbon;

            public RibbonCacheListener(MainRibbon mainRibbon)
            {
                this.mainRibbon = mainRibbon;
            }

            void ShapePersistence.CacheListener.OnCacheRenewed()
            {
                this.redrawRibbon();
            }

            void ShapePersistence.CacheListener.OnItemAdded(ShapeFavorite addedItem)
            {
                this.redrawRibbon();
            }

            void ShapePersistence.CacheListener.OnItemRemoved(ShapeFavorite removedItem)
            {
                this.redrawRibbon();
            }

            private void redrawRibbon()
            {
                mainRibbon.ribbon.Invalidate();
            }
        }

        public bool IsAdmin(Office.IRibbonControl control)
        {
            return BuildEnvironment.IsAdminBuild();
        }

        public void OnOpenSharedFavButton(Office.IRibbonControl control)
        {
            logger.Log("Open Button pressed.");
            SharedFavView.ShowOrFocus();
        }

        public void OnInfoButton(Office.IRibbonControl control)
        {
            logger.Log("Info Button pressed.");
            System.Diagnostics.Process.Start(Constants.GITHUB_URL);
        }

        public void OnTutorialButton(Office.IRibbonControl control)
        {
            logger.Log("Tutorial Button pressed.");
            System.Diagnostics.Process.Start(Constants.TUTORIAL_URL);
        }

        public void OnStoryButton(Office.IRibbonControl control)
        {
            logger.Log("Story Button pressed.");
            System.Diagnostics.Process.Start(Constants.STORY_URL);
        }

        public void OnInstallDefaultThemeButton(Office.IRibbonControl control)
        {
            logger.Log("Install default theme pressed.");
            TryInstallDefaultTheme();
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
                    //import is done. Maybe set theme as default?
                    if(shapePersistence.GetPersistedTheme() == null)
                    {
                        //no theme to import
                        MessageBox.Show("Successfully imported favorites.");
                    }
                    else
                    {
                        HandlePersistedThemeImport();
                    }
                }
                else
                {
                    MessageBox.Show("An error occured while importing favorites.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public void OnDeletePresentationStore(Office.IRibbonControl control)
        {
            logger.Log("Delete Presentation button clicked.");
            shapePersistence.DeletePresentationStore();
        }

        public void OnAddPresentation(Office.IRibbonControl control)
        {
            logger.Log("Add Presentation button clicked.");
            var filePath = DialogUtil.GetFilePathViaDialog(isSaveAction: false, filter: DialogUtil.POWERPOINT_PRESENTATION_FILTER);
            if(filePath == null)
            {
                return;
            }
            if(!filePath.EndsWith(".pptx"))
            {
                MessageBox.Show("This file format is not supported.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if(ShapePersistence.INSTANCE.SavePresentation(filePath))
            {
                MessageBox.Show("Presentation was successfully imported.", "Success", MessageBoxButtons.OK);
            }
            else
            {
                MessageBox.Show("An error occured while importing the presentation.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void HandlePersistedThemeImport()
        {
            if (MessageBox.Show("Favorites were successfully imported. Do you want to install the imported theme as default?",
                                     "Install theme",
                                     MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                logger.Log("User wants to install default theme.");
                TryInstallDefaultTheme();   
            }
        }

        private void TryInstallDefaultTheme()
        {
            logger.Log("Trying to install default theme.");
            if (shapePersistence.InstallPersistedDefaultTheme())
            {
                MessageBox.Show("Successfully installed default theme.");
            }
            else
            {
                MessageBox.Show("An error occured while installing default theme.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnExportSharedFavButton(Office.IRibbonControl control)
        {
            logger.Log("Export button clicked.");
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
            logger.Log("Copy from Clipboard clicked.");
            if (!this.shapePersistence.SaveShapeFromClipBoard(null))
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
            }
            else
            {
                logger.Log("Could not save selection " + selection);
            }
        }

        public string GetVersionLabel(Office.IRibbonControl control)
        {
            return "Version " + BuildEnvironment.GetFullVersion();
        }

        public int GetItemCount(Office.IRibbonControl control)
        {
            return this.shapeService.GetShapesByTypes(GetShapeTypesForControl(control)).Count;
        }

        public int GetPresentationCount(Office.IRibbonControl control)
        {
            return shapePersistence.GetPresentationStoreSlideCount();
        }

        public Bitmap GetPresentationImage(Office.IRibbonControl control, int index)
        {
            return GetScaledImage(shapePersistence.GetPresentationStoreSlideThumbByIndex(index));
        }

        public void OnPresentationAction(Office.IRibbonControl control, string id, int index)
        {
            shapePersistence.PastePresentationStoreSlide(index);
        }


        private IEnumerable<Office.MsoShapeType> GetShapeTypesForControl(Office.IRibbonControl control)
        {
            switch (control.Id)
            {
                case "sharedFavsCharts":
                    return SupportedShapeTypes.Charts;
                case "sharedFavsTables":
                    return SupportedShapeTypes.Tables;
                case "sharedFavsShapes":
                    return SupportedShapeTypes.Shapes;
                case "sharedFavsPictures":
                    return SupportedShapeTypes.Pictures;
            }
            throw new ArgumentOutOfRangeException("Unknown shape type");
        }

        public Bitmap GetItemImage(Office.IRibbonControl control, int index)
        {
            var thumbnail = this.shapeService.GetShapesByTypes(GetShapeTypesForControl(control))[index].Thumbnail;
            return GetScaledImage(thumbnail);
        }

        private Bitmap GetScaledImage(Image original)
        {
            float targetWidth = ITEM_WIDTH;
            float targetHeight = ITEM_HEIGHT;
            float scale = Math.Min(targetWidth / original.Width, targetHeight / original.Height);
            logger.Log("Using scale factor: " + scale);
            var scaledWidth = (int)(original.Width * scale);
            var scaledHeight = (int)(original.Height * scale);
            var brush = new SolidBrush(Color.White);
            var result = new Bitmap((int)targetWidth, (int)targetHeight);
            using (var graphics = Graphics.FromImage(result))
            {
                graphics.FillRectangle(brush, new RectangleF(0, 0, targetWidth, targetHeight));
                graphics.DrawImage(original, new Rectangle(((int)targetWidth - scaledWidth) / 2, ((int)targetHeight - scaledHeight) / 2, scaledWidth, scaledHeight));
            }
            return result;
        }

        public void OnItemAction(Office.IRibbonControl control, string id, int index)
        {
            var shapeTypes = GetShapeTypesForControl(control);
            logger.Log("Chart of type " + shapeTypes + " with id " + index + " clicked.");
            this.shapeService.PasteToCurrentPresentation(this.shapeService.GetShapesByTypes(shapeTypes)[index]);
        }

        public Bitmap GetPentagonImage(Office.IRibbonControl control)
        {
            return new Bitmap(Properties.Resources.Pentagon);
        }

        public Bitmap GetTutorialImage(Office.IRibbonControl control)
        {
            return new Bitmap(Properties.Resources.Tutorial);
        }

        public Bitmap GetInfoImage(Office.IRibbonControl control)
        {
            return new Bitmap(Properties.Resources.Info);
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
