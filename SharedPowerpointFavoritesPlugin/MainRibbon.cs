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
using SharedPowerpointFavoritesPlugin.view;
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
            new InfoDialog().ShowDialog();
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
        
        public int GetItemCountCharts(Office.IRibbonControl control)
        {
            return this.GetItemCount(Office.MsoShapeType.msoChart);
        }

        public System.Drawing.Bitmap GetItemImageCharts(Office.IRibbonControl control, int index)
        {
            return this.GetItemImage(Office.MsoShapeType.msoChart, index);
        }

        public void OnChartAction(Office.IRibbonControl control, string id, int index)
        {
            this.OnAction(Office.MsoShapeType.msoChart, index);
        }
        
        public int GetItemCountTables(Office.IRibbonControl control)
        {
            return this.GetItemCount(Office.MsoShapeType.msoTable);
        }

        public System.Drawing.Bitmap GetItemImageTables(Office.IRibbonControl control, int index)
        {
            return this.GetItemImage(Office.MsoShapeType.msoTable, index);
        }

        public void OnTableAction(Office.IRibbonControl control, string id, int index)
        {
            this.OnAction(Office.MsoShapeType.msoTable, index);
        }

        public int GetItemCountShapes(Office.IRibbonControl control)
        {
            return this.GetItemCount(Office.MsoShapeType.msoAutoShape);
        }

        public System.Drawing.Bitmap GetItemImageShapes(Office.IRibbonControl control, int index)
        {
            return this.GetItemImage(Office.MsoShapeType.msoAutoShape, index);
        }

        public void OnShapeAction(Office.IRibbonControl control, string id, int index)
        {
            this.OnAction(Office.MsoShapeType.msoAutoShape, index);
        }

        public int GetItemCountPictures(Office.IRibbonControl control)
        {
            return this.GetItemCount(Office.MsoShapeType.msoPicture);
        }

        public System.Drawing.Bitmap GetItemImagePictures(Office.IRibbonControl control, int index)
        {
            return this.GetItemImage(Office.MsoShapeType.msoPicture, index);
        }

        public void OnPictureAction(Office.IRibbonControl control, string id, int index)
        {
            this.OnAction(Office.MsoShapeType.msoPicture, index);
        }

        public int GetItemCountGroups(Office.IRibbonControl control)
        {
            return this.GetItemCount(Office.MsoShapeType.msoGroup);
        }

        public System.Drawing.Bitmap GetItemImageGroups(Office.IRibbonControl control, int index)
        {
            return this.GetItemImage(Office.MsoShapeType.msoGroup, index);
        }

        public void OnGroupAction(Office.IRibbonControl control, string id, int index)
        {
            this.OnAction(Office.MsoShapeType.msoGroup, index);
        }


        private int GetItemCount(Office.MsoShapeType shapeType)
        {
            return this.shapeService.GetShapesByType(shapeType).Count;
        }

        private System.Drawing.Bitmap GetItemImage(Office.MsoShapeType shapeType, int index)
        {
            var thumbnail = this.shapeService.GetShapesByType(shapeType)[index].Thumbnail;
            return new System.Drawing.Bitmap(thumbnail, 100, GetHeight(100, thumbnail));
        }
        
        private int GetHeight(int width, Image image)
        {
            var ratio = (float)image.Height / (float)image.Width;
            var calculatedHeight = ratio * width;
            logger.Log("Using bitmap size: " + calculatedHeight + "|" + width);
            return (int) calculatedHeight;
        }

        private void OnAction(Office.MsoShapeType shapeType, int index)
        {
            logger.Log("Chart of type " + shapeType + " with id " + index + " clicked.");
            this.shapeService.PasteToCurrentPresentation(this.shapeService.GetShapesByType(shapeType)[index]);
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
