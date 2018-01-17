﻿using SharedPowerpointFavoritesPlugin.model;
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
            if (util.BuildEnvironment.IsAdminBuild())
            {
                pictureBox.ContextMenu = GetContextMenu(pictureBox);
            }
            return pictureBox;
        }

        private ContextMenu GetContextMenu(PictureBox pictureBox)
        {
            var contextMenu = new ContextMenu();
            var deleteItem = new MenuItem("Delete...", new EventHandler((sender, args) => HandleDeleteItem(pictureBox)));
            contextMenu.MenuItems.Add(deleteItem);
            return contextMenu;
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

        private void HandlePictureBoxDoubleClick(ShapeFavorite shape)
        {
            logger.Log("Double click on picture box. Pasting shape.");
            this.shapeService.PasteToCurrentPresentation(shape);
        }

        private void SharedFavView_Load(object sender, EventArgs e)
        {
            this.InitializeTabPages();
            this.RedrawFavorites();
            var updateListener = new UpdateFavViewListener();
            this.shapePersistance.RegisterCacheListener(updateListener);
            this.FormClosed += new FormClosedEventHandler((_sender, _args) =>
            {
                shapePersistance.RemoveCacheListener(updateListener);
            });
        }

        private void InitializeTabPages()
        {
            this.CreateTabPage(Office.MsoShapeType.msoChart, "Charts");
            this.CreateTabPage(Office.MsoShapeType.msoAutoShape, "Auto Shapes");
            this.CreateTabPage(Office.MsoShapeType.msoTable, "Tables");
            //add further pages here...
        }

        private void CreateTabPage(Office.MsoShapeType shapeType, string caption)
        {
            var parentControl = this.tabControl1;
            var tabPage = new TabPage();
            tabPage.Width = parentControl.Width - 8;
            tabPage.Height = parentControl.Height - 28;
            var panel = this.GetPanel(tabPage);
            tabPage.Controls.Add(panel);
            this.panels.Add(shapeType, panel);
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
            public void OnCacheRenewed()
            {
                logger.Log("CacheRenewedListener fired.");
                SharedFavView.CURRENT_INSTANCE.RedrawFavorites();
            }

            public void OnItemAdded(ShapeFavorite addedItem)
            {
                logger.Log("ItemAddedListener fired.");
                SharedFavView.CURRENT_INSTANCE.DrawShape(addedItem, SharedFavView.CURRENT_INSTANCE.panels[addedItem.Shape.Type]);
            }

            public void OnItemRemoved(ShapeFavorite removedItem)
            {
                logger.Log("ItemRemovedListener fired.");
                SharedFavView.CURRENT_INSTANCE.RedrawFavorites(); //it would be difficult to do this incrementally since that would imply reordering of the picture boxes...
            }
        }
    }
}
