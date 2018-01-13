using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Core = Microsoft.Office.Core;
using Interop = Microsoft.Office.Interop;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Path = System.IO.Path;
using Directory = System.IO.Directory;
using SharedPowerpointFavoritesPlugin.model;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SharedPowerpointFavoritesPlugin
{
    class ShapePersistence
    {
        public static ShapePersistence INSTANCE = new ShapePersistence();
        public const string PERSISTENCE_DIR = ".sharedpowerpointfavorites";
        public const string PERSISTENCE_EXTENSION = ".pptx";
        public const string PNG_EXTENSION = ".png";
        private List<CacheListener> cacheListeners = new List<CacheListener>();
        private List<ShapeFavorite> _cachedShapes; //backing dictionary
        private List<ShapeFavorite> CachedShapes
        {
            get
            {
                if (_cachedShapes == null)
                {
                    this.LoadShapes();
                }
                return this._cachedShapes;
            }
            set
            {
                _cachedShapes = value;
                InformCacheListenersOnRenew();
            }
        }

        private void InformCacheListenersOnRenew()
        {
            foreach(CacheListener listener in cacheListeners)
            {
                listener.onCacheRenewed();
            }
        }

        private void InformCacheListenersOnItemAdded(ShapeFavorite addedItem)
        {
            foreach(CacheListener listener in cacheListeners)
            {
                listener.onItemAdded(addedItem);
            }
        }

        internal string GetThumbnail(ShapeFavorite shape)
        {
            var thumbnailPath = GetThumbnailPath(shape);
            if (!System.IO.File.Exists(thumbnailPath))
            {
                DebugLogger.Log("Thumbnail does not exist. Creating one.");
                var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                var temporaryPresentation = Globals.ThisAddIn.Application.Presentations.Open(shape.FilePath, Core.MsoTriState.msoTrue, Core.MsoTriState.msoTrue, Core.MsoTriState.msoFalse);
                var targetSlide = temporaryPresentation.Slides[1];
                var targetShape = targetSlide.Shapes[1];
                var shapeExportArgs = new object[] { thumbnailPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG, 0, 0, PowerPoint.PpExportMode.ppRelativeToSlide };
                targetShape.GetType().InvokeMember("Export", System.Reflection.BindingFlags.InvokeMethod, null, targetShape, shapeExportArgs); //ATTENTION. This is the risky part...
                temporaryPresentation.Close();
                stopwatch.Stop();
                DebugLogger.Log("Creating thumbnail took " + stopwatch.ElapsedMilliseconds);
            }
            return thumbnailPath;
        }

        internal void RemoveCacheListener(CacheListener updateListener)
        {
            DebugLogger.Log("Removing cache listener " + updateListener);
            this.cacheListeners.Remove(updateListener);
        }

        private ShapePersistence()
        {
            //singleton
        }

        private string GetThumbnailPath(ShapeFavorite shape)
        {
            return shape.FilePath.Replace(PERSISTENCE_EXTENSION, PNG_EXTENSION);
        }

        public void SaveShapeFromClipBoard()
        {
            var temporaryPresentation = Globals.ThisAddIn.Application.Presentations.Add(Core.MsoTriState.msoFalse);
            var targetSlide = temporaryPresentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            var newUuid = Guid.NewGuid().ToString();
            var fileName = GetFileName(newUuid);
            var persistenceFile = GetPersistenceFile(fileName);
            targetSlide.Shapes.Paste();
            DebugLogger.Log("Saving shape.");
            var cachedShapes = CachedShapes; // ensure this is already loaded before saving!
            var shapeToSave = new ShapeFavorite(persistenceFile, targetSlide.Shapes[1]);
            temporaryPresentation.SaveAs(shapeToSave.FilePath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Core.MsoTriState.msoFalse);
            temporaryPresentation.Close();
            this.GetThumbnail(shapeToSave); //create thumbnail
            cachedShapes.Add(shapeToSave);
            InformCacheListenersOnItemAdded(shapeToSave);
        }

        public void SaveShape(Shape shape)
        {
            shape.Copy();
            this.SaveShapeFromClipBoard();
        }


        public List<ShapeFavorite> GetShapes()
        {
            return new List<ShapeFavorite>(CachedShapes);
        }


        internal void LoadShapes()
        {
            var persistenceDir = GetPersistenceDir();
            DebugLogger.Log("Loading shapes from persistence directory: " + persistenceDir);
            string[] filePaths = Directory.GetFiles(@persistenceDir, "*" + PERSISTENCE_EXTENSION,
                                         System.IO.SearchOption.TopDirectoryOnly);
            var loadedShapes = new List<ShapeFavorite>();
            foreach (string file in filePaths)
            {
                DebugLogger.Log("Reading file " + file);
                List<Shape> shapesFromFile = this.GetShapesFromFile(file);
                foreach (Shape shape in shapesFromFile)
                {
                    loadedShapes.Add(new ShapeFavorite(file, shape));
                    DebugLogger.Log("Loaded shape from file " + file);
                }
            }
            CachedShapes = loadedShapes;
        }


        private List<Shape> GetShapesFromFile(string file)
        {
            var temporaryPresentation = Globals.ThisAddIn.Application.Presentations.Open(file, Core.MsoTriState.msoTrue, Core.MsoTriState.msoTrue, Core.MsoTriState.msoFalse);
            List<Shape> result = new List<Shape>();
            foreach (Shape shape in temporaryPresentation.Slides[1].Shapes.Cast<Shape>().ToList())
            {
                result.Add(shape);
            }
            return result;
        }


        private string GetFileName(string uuid)
        {
            return uuid + PERSISTENCE_EXTENSION;
        }

        private string GetPersistenceFile(string fileName)
        {
            var separator = Path.DirectorySeparatorChar;
            var fileDir = GetPersistenceDir();
            var filePath = fileDir + separator + fileName;
            DebugLogger.Log("Using file path: " + filePath);
            return filePath;
        }

        internal string GetPersistenceDir()
        {
            var homePath = Environment.ExpandEnvironmentVariables("%HOMEDRIVE%%HOMEPATH%");
            var separator = Path.DirectorySeparatorChar;
            var persistenceDir = homePath + separator + PERSISTENCE_DIR;
            Directory.CreateDirectory(persistenceDir);
            return persistenceDir;
        }

        internal void RegisterCacheListener(CacheListener listener)
        {
            DebugLogger.Log("Adding cache listener: " + listener);
            this.cacheListeners.Add(listener);
        }

        internal interface CacheListener
        {
            void onCacheRenewed();

            void onItemAdded(ShapeFavorite addedItem);
        }
    }
}
