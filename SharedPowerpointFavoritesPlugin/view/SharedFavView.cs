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

using SharedPowerpointFavoritesPlugin.model;
using SharedPowerpointFavoritesPlugin.util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SharedPowerpointFavoritesPlugin
{
    public partial class SharedFavView : Form
    {

        private static SharedFavView CURRENT_INSTANCE;
        private ShapePersistence shapePersistance = ShapePersistence.INSTANCE;
        private ShapeService shapeService = ShapeService.INSTANCE;
        private Dictionary<PictureBox, ShapeFavorite> displayedShapes = new Dictionary<PictureBox, ShapeFavorite>();
        private ImportExportService importExportService = ImportExportService.INSTANCE;
        private static readonly DebugLogger logger = DebugLogger.GetLogger(typeof(SharedFavView).Name);
        private Dictionary<Office.MsoShapeType, Panel> panels = new Dictionary<Office.MsoShapeType, Panel>();

        public SharedFavView()
        {
            InitializeComponent();
        }

        private void SharedFavView_FormClosed(object sender, FormClosedEventArgs e)
        {
            CURRENT_INSTANCE = null;
        }

        private void SharedFavView_Shown(object sender, EventArgs e)
        {
            CURRENT_INSTANCE = this;
        }

        public static void ShowOrFocus()
        {
            if (CURRENT_INSTANCE == null)
            {
                new SharedFavView().Show();
            }
            else
            {
                CURRENT_INSTANCE.BringToFront();
            }
        }

        private void DrawShape(ShapeFavorite shape, Panel panel)
        {
            var index = panel.Controls.Count;
            var pictureBox = GetPictureBox(shape, index);
            panel.Controls.Add(pictureBox);
            this.displayedShapes.Add(pictureBox, shape);
        }

        private PictureBox GetPictureBox(ShapeFavorite shape, int index)
        {
            var pictureBox = new PictureBox();
            pictureBox.Height = 100;
            pictureBox.Width = 100;
            int column = index % 2;
            int row = (index - (index % 2)) / 2;
            pictureBox.Location = new Point(4 + (column * 120), 4 + (row * 120));
            pictureBox.Image = shape.Thumbnail;
            pictureBox.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox.MouseDoubleClick += new MouseEventHandler((sender, args) => HandlePictureBoxDoubleClick(shape));
            pictureBox.ContextMenu = GetContextMenu(pictureBox);
            return pictureBox;
        }

        private ContextMenu GetContextMenu(PictureBox pictureBox)
        {
            var contextMenu = new ContextMenu();
            var addItem = new MenuItem("Add to Slide", new EventHandler((sender, args) => HandleAddItemToSlide(pictureBox)));
            contextMenu.MenuItems.Add(addItem);
            if (BuildEnvironment.IsAdminBuild())
            {
                var deleteItem = new MenuItem("Delete...", new EventHandler((sender, args) => HandleDeleteItem(pictureBox)));
                contextMenu.MenuItems.Add(deleteItem);
                var moveUpItem = new MenuItem("Move up", new EventHandler((sender, args) => HandleMoveUpItem(pictureBox)));
                contextMenu.MenuItems.Add(moveUpItem);
                var moveToTopItem = new MenuItem("Move to Top", new EventHandler((sender, args) => HandleMoveToTopItem(pictureBox)));
                contextMenu.MenuItems.Add(moveToTopItem);
            }
            return contextMenu;
        }

        private void HandleAddItemToSlide(PictureBox pictureBox)
        {
            logger.Log("User clicked add item to slide.");
            this.shapeService.PasteToCurrentPresentation(this.displayedShapes[pictureBox]);
        }

        private void HandleDeleteItem(PictureBox pictureBox)
        {
            logger.Log("User clicked delete.");
            if (!this.AskForDeleteConfirmation())
            {
                logger.Log("User cancelled item deletion.");
                return;
            }
            this.shapeService.DeleteShape(this.displayedShapes[pictureBox]);
        }

        private void HandleMoveUpItem(PictureBox pictureBox)
        {
            logger.Log("User clicked moveUp.");
            this.shapeService.MoveUpShape(this.displayedShapes[pictureBox]);
        }

        private void HandleMoveToTopItem(PictureBox pictureBox)
        {
            logger.Log("User clicked moveToTop.");
            this.shapeService.MoveShapeToTop(this.displayedShapes[pictureBox]);
        }

        private void HandlePictureBoxDoubleClick(ShapeFavorite shape)
        {
            logger.Log("Double click on picture box. Pasting shape.");
            this.shapeService.PasteToCurrentPresentation(shape);
        }

        private void SharedFavView_Load(object sender, EventArgs e)
        {
            this.InitializeTabPages();
            this.RedrawFavorites();
            var updateListener = new UpdateFavViewListener(this);
            this.shapePersistance.RegisterCacheListener(updateListener);
            this.FormClosed += new FormClosedEventHandler((_sender, _args) =>
            {
                shapePersistance.RemoveCacheListener(updateListener);
            });
        }

        private void InitializeTabPages()
        {
            foreach (KeyValuePair<string, List<Office.MsoShapeType>> entry in SupportedShapeTypes.All)
            {
                this.CreateTabPage(entry.Key, entry.Value.ToArray());
            }
        }

        private void CreateTabPage(string caption, params Office.MsoShapeType[] shapeTypes)
        {
            var parentControl = this.tabControl1;
            var tabPage = new TabPage();
            tabPage.Width = parentControl.Width - 8;
            tabPage.Height = parentControl.Height - 28;
            var panel = this.GetPanel(tabPage);
            tabPage.Controls.Add(panel);
            foreach (Office.MsoShapeType shapeType in shapeTypes)
            {
                this.panels.Add(shapeType, panel);
            }
            tabPage.Text = caption;
            parentControl.Controls.Add(tabPage);
        }

        private Panel GetPanel(TabPage tabPage)
        {
            var panel = new Panel();
            panel.Width = tabPage.Width;
            panel.Height = tabPage.Height;
            panel.AutoScroll = true;
            return panel;
        }

        private void RedrawFavorites()
        {
            logger.Log("Reloading all favorites.");
            this.RemoveAllPictureBoxes();
            foreach (Office.MsoShapeType shapeType in this.panels.Keys)
            {
                List<ShapeFavorite> shapes = this.shapeService.GetShapesByType(shapeType);
                foreach (ShapeFavorite shape in shapes)
                {
                    this.DrawShape(shape, this.panels[shapeType]);
                }

            }
        }

        private void RemoveAllPictureBoxes()
        {
            foreach (PictureBox pictureBox in displayedShapes.Keys)
            {
                foreach (Panel panel in this.panels.Values)
                {
                    if (panel.Controls.Contains(pictureBox))
                    {
                        panel.Controls.Remove(pictureBox);
                    }
                }
            }
            this.displayedShapes.Clear();
        }

        private bool AskForDeleteConfirmation()
        {
            return DialogUtil.AskForConfirmation("Are you sure you want to delete this item?");
        }

        class UpdateFavViewListener : ShapePersistence.CacheListener
        {
            private SharedFavView sharedFavView;

            public UpdateFavViewListener(SharedFavView sharedFavView)
            {
                this.sharedFavView = sharedFavView;
            }

            public void OnCacheRenewed()
            {
                logger.Log("CacheRenewedListener fired.");
                this.sharedFavView.RedrawFavorites();
            }

            public void OnItemAdded(ShapeFavorite addedItem)
            {
                logger.Log("ItemAddedListener fired.");
                this.sharedFavView.DrawShape(addedItem, this.sharedFavView.panels[addedItem.Shape.Type]);
            }

            public void OnItemRemoved(ShapeFavorite removedItem)
            {
                logger.Log("ItemRemovedListener fired.");
                this.sharedFavView.RedrawFavorites(); //it would be difficult to do this incrementally since that would imply reordering of the picture boxes...
            }
        }
    }
}
