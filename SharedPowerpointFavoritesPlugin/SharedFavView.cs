using SharedPowerpointFavoritesPlugin.model;
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

        private void DrawShape(ShapeFavorite shape)
        {
            var index = this.displayedShapes.Count;
            var pictureBox = GetPictureBox(shape, index);
            this.panel1.Controls.Add(pictureBox);
            this.displayedShapes.Add(pictureBox, shape);
        }

        private PictureBox GetPictureBox(ShapeFavorite shape, int index)
        {
            var pictureBox = new PictureBox();
            pictureBox.Height = 100;
            pictureBox.Width = 100;
            pictureBox.Location = new Point(4, 4 + index * 120);
            pictureBox.Image = shape.Thumbnail;
            pictureBox.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox.MouseDoubleClick += new MouseEventHandler((sender, args) => HandlePictureBoxDoubleClick(shape));
            return pictureBox;
        }

        private void HandlePictureBoxDoubleClick(ShapeFavorite shape)
        {
            this.shapeService.PasteToCurrentPresentation(shape);
        }

        private void SharedFavView_Load(object sender, EventArgs e)
        {
            this.ReloadFavorites();
            var updateListener = new UpdateFavViewListener();
            this.shapePersistance.RegisterCacheListener(updateListener);
            this.FormClosed += new FormClosedEventHandler((_sender, _args) => {
                shapePersistance.RemoveCacheListener(updateListener);
            });
        }

        private void ReloadFavorites()
        {
            DebugLogger.Log("Reloading all favorites.");
            //TODO remove all existing pictureBoxes
            foreach(PictureBox pictureBox in displayedShapes.Keys)
            {
                this.panel1.Controls.Remove(pictureBox);
            }
            this.displayedShapes.Clear();
            List<ShapeFavorite> shapes = this.shapePersistance.GetShapes();
            foreach (ShapeFavorite shape in shapes)
            {
                this.DrawShape(shape);
            }
        }

        private void saveShapeButton_Click(object sender, EventArgs e)
        {
            this.shapePersistance.SaveShapeFromClipBoard();
        }

        private void importButton_Click(object sender, EventArgs e)
        {
            var filePath = GetFilePathViaDialog(isSaveAction : false);
            if (filePath != null)
            {
                if(this.importExportService.ImportFromFile(filePath))
                {
                    MessageBox.Show("Successfully imported favorites.");
                }
                else
                {
                    MessageBox.Show("An error occured while importing favorites.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void exportButton_Click(object sender, EventArgs e)
        {
            var filePath = GetFilePathViaDialog(isSaveAction : true);
            if (filePath != null)
            {
                if(this.importExportService.ExportToFile(filePath))
                {
                    MessageBox.Show("Successfully exported favorites.");
                }
                else
                {
                    MessageBox.Show("An error occured while exporting favorites.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private string GetFilePathViaDialog(bool isSaveAction)
        {
            FileDialog openFileDialog = isSaveAction ? (new SaveFileDialog() as FileDialog) : (new OpenFileDialog() as FileDialog);
            openFileDialog.InitialDirectory = Environment.ExpandEnvironmentVariables("%HOMEDRIVE%%HOMEPATH%");
            openFileDialog.Filter = "SharedPowerpointFavorites (*.zip)|*.zip";
            openFileDialog.RestoreDirectory = true;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog.FileName;
            }
            else
            {
                DebugLogger.Log("No file chosen.");
                return null;
            }
        }

        class UpdateFavViewListener : ShapePersistence.CacheListener
        {
            public void onCacheRenewed()
            {
                DebugLogger.Log("CacheRenewedListener fired.");
                SharedFavView.CURRENT_INSTANCE.ReloadFavorites();
            }

            public void onItemAdded(ShapeFavorite addedItem)
            {
                DebugLogger.Log("ItemAddedListener fired.");
                SharedFavView.CURRENT_INSTANCE.DrawShape(addedItem);
            }
        }
    }
}
