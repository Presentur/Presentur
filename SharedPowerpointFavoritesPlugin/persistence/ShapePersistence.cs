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
using FileInfo = System.IO.FileInfo;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.Drawing;

namespace SharedPowerpointFavoritesPlugin
{
    class ShapePersistence
    {
        private static readonly string STRUCTURE_PERSISTANCE_FILE = GetPersistenceDir() + Path.DirectorySeparatorChar + "structure.json";
        private const string DEFAULT_THEME_FILE = "Default Theme.thmx";
        public const string PERSISTENCE_DIR = ".sharedpowerpointfavorites";
        private const string PRESENTATION_DIR = "presentationstore";
        public const string PERSISTENCE_EXTENSION = ".pptx";
        public const string PNG_EXTENSION = ".png";
        private const string THEME_EXTENSION = ".thmx";
        private const string STORED_PRESENTATION_FILENAME = "presentation.pptx";

        private static readonly DebugLogger logger = DebugLogger.GetLogger(typeof(ShapePersistence).Name);
        public static ShapePersistence INSTANCE = new ShapePersistence();
        private List<CacheListener> cacheListeners = new List<CacheListener>();
        private List<ShapeFavorite> _cachedShapes; //backing list
        

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

        internal void SetShapeTheme(string themeName)
        {
            this.RemoveAllSavedThemes();
            logger.Log("Saving current theme.");
            var dir = GetPersistenceDir();
            var file = dir + Path.DirectorySeparatorChar + themeName + THEME_EXTENSION;
            Globals.ThisAddIn.Application.ActivePresentation.SaveCopyAs(file);
        }

        internal void InstallPersistedTheme()
        {
            var persistedTheme = this.GetPersistedTheme();
            if (persistedTheme == null)
            {
                logger.Log("No theme found to install.");
                return;
            }
            WriteTheme(persistedTheme, Path.GetFileName(persistedTheme));
        }

        internal bool InstallPersistedDefaultTheme()
        {
            var persistedTheme = this.GetPersistedTheme();
            if (persistedTheme == null)
            {
                logger.Log("No theme found to install as default.");
                return false;
            }
            logger.Log("Writing default theme.");
            WriteTheme(persistedTheme, DEFAULT_THEME_FILE);
            return true;
        }

        private void WriteTheme(string source, string targetFileName)
        {
            logger.Log("Installing theme: " + source);
            var appDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var themeDir = "Microsoft\\Templates\\Document Themes";
            var targetDir = Path.Combine(appDataDir, themeDir);
            var targetFile = Path.Combine(targetDir, targetFileName);
            logger.Log("Target file is: " + targetFile);
            File.Copy(source, targetFile, true); //overwrites existing theme if present
        }

        internal string GetPersistedTheme()
        {
            var themes = Directory.GetFiles(GetPersistenceDir(), "*" + THEME_EXTENSION,
                                         System.IO.SearchOption.TopDirectoryOnly);
            if (themes.Length > 0)
            {
                return themes[0];
            }
            return null;
        }

        private void RemoveAllSavedThemes()
        {
            logger.Log("Removing all saved themes.");
            this.DeleteIfExtant(Directory.GetFiles(GetPersistenceDir(), "*" + THEME_EXTENSION,
                                         System.IO.SearchOption.TopDirectoryOnly));
        }

        private void InformCacheListenersOnRenew()
        {
            foreach (CacheListener listener in cacheListeners)
            {
                listener.OnCacheRenewed();
            }
        }

        private void InformCacheListenersOnItemAdded(ShapeFavorite addedItem)
        {
            foreach (CacheListener listener in cacheListeners)
            {
                listener.OnItemAdded(addedItem);
            }
        }

        private void InformCacheListenersOnItemRemoved(ShapeFavorite removedItem)
        {
            foreach (CacheListener listener in cacheListeners)
            {
                listener.OnItemRemoved(removedItem);
            }
        }

        internal string GetThumbnail(ShapeFavorite shape)
        {
            var thumbnailPath = GetThumbnailPath(shape);
            if (!System.IO.File.Exists(thumbnailPath))
            {
                logger.Log("Thumbnail does not exist. Creating one.");
                var temporaryPresentation = Globals.ThisAddIn.Application.Presentations.Open(shape.FilePath, Core.MsoTriState.msoTrue, Core.MsoTriState.msoTrue, Core.MsoTriState.msoFalse);
                var targetSlide = temporaryPresentation.Slides[1];
                var targetShape = targetSlide.Shapes[1];
                CreateThumbnail(thumbnailPath, targetShape);
                temporaryPresentation.Close();
            }
            return thumbnailPath;
        }

        internal void DeletePresentationStore()
        {
            var presentationDir = GetPresentationDir();
            Directory.Delete(presentationDir, true);
            InformCacheListenersOnRenew();
        }

        internal bool SavePresentation(string filePath)
        {
            DeletePresentationStore(); //we currently only support one stored presentation at a time
            var newUuid = Guid.NewGuid().ToString();
            var presentationDir = GetPresentationDir();
            var targetDir = presentationDir + Path.DirectorySeparatorChar + newUuid;
            var targetPresentationFile = targetDir + Path.DirectorySeparatorChar + STORED_PRESENTATION_FILENAME;
            Directory.CreateDirectory(targetDir);
            File.Copy(filePath, targetPresentationFile);
            if (!CreatePresentationThumbnails(targetPresentationFile))
            {
                return false;
            }
            InformCacheListenersOnRenew();
            return true;
        }

        private bool CreatePresentationThumbnails(string presentationFile)
        {
            var temporaryPresentation = Globals.ThisAddIn.Application.Presentations.Open(presentationFile, Core.MsoTriState.msoTrue, Core.MsoTriState.msoTrue, Core.MsoTriState.msoFalse);
            temporaryPresentation.SaveAs(Directory.GetParent(presentationFile).FullName, PowerPoint.PpSaveAsFileType.ppSaveAsPNG);
            temporaryPresentation.Close();
            return true;
        }

        private void CreateThumbnail(string thumbnailPath, Shape targetShape)
        {
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            var shapeExportArgs = new object[] { thumbnailPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG, 0, 0, PowerPoint.PpExportMode.ppRelativeToSlide };
            targetShape.GetType().InvokeMember("Export", System.Reflection.BindingFlags.InvokeMethod, null, targetShape, shapeExportArgs); //ATTENTION. This is the risky part...
            stopwatch.Stop();
            logger.Log("Creating thumbnail took " + stopwatch.ElapsedMilliseconds);
        }

        internal void DeleteShape(ShapeFavorite shapeFavorite)
        {
            logger.Log("Deleting shape of type " + shapeFavorite.Shape.Type);
            var thumbnail = GetThumbnailPath(shapeFavorite);
            this.DeleteIfExtant(shapeFavorite.FilePath, thumbnail);
            CachedShapes.Remove(shapeFavorite);
            this.InformCacheListenersOnItemRemoved(shapeFavorite);
        }

        private void DeleteIfExtant(params string[] paths)
        {
            foreach (string path in paths)
            {
                if (File.Exists(path))
                {
                    logger.Log("Deleting file " + path);
                    try
                    {
                        File.Delete(path);
                    }
                    catch (Exception e)
                    {
                        logger.Log("Exception while deleting file " + path + ". Exception is: " + e.Message);
                    }
                }
                else
                {
                    logger.Log("File does not exist: " + path);
                }
            }
        }

        internal void RemoveCacheListener(CacheListener updateListener)
        {
            logger.Log("Removing cache listener " + updateListener);
            this.cacheListeners.Remove(updateListener);
        }

        private ShapePersistence()
        {
            //singleton
            this.cacheListeners.Add(new StructurePersistanceListener(this, STRUCTURE_PERSISTANCE_FILE));
        }

        private string GetThumbnailPath(ShapeFavorite shape)
        {
            return shape.FilePath.Replace(PERSISTENCE_EXTENSION, PNG_EXTENSION);
        }

        public bool MoveUp(ShapeFavorite shapeFavorite, bool toTop)
        {
            logger.Log("Moving up shapeFavorite: " + shapeFavorite);
            var shapeFavorites = new List<ShapeFavorite>(CachedShapes);
            int targetIndex = -1;
            for (int i = 0; i < shapeFavorites.Count; i++)
            {
                if (IsPeerShapeType(shapeFavorites[i].Shape.Type, shapeFavorite.Shape.Type) && !(shapeFavorites[i].Equals(shapeFavorite)))
                {
                    targetIndex = i;
                    if (toTop)
                    {
                        break;
                    }
                }
                if (shapeFavorites[i].Equals(shapeFavorite))
                {
                    break;
                }
            }
            if (targetIndex == -1)
            {
                logger.Log("Did not find any shape with lower index. Not moving up.");
                return false;
            }
            logger.Log("Setting targetIndex for moving up: " + targetIndex);
            shapeFavorites.Remove(shapeFavorite);
            shapeFavorites.Insert(targetIndex, shapeFavorite);
            CachedShapes = shapeFavorites;
            return true;
        }

        internal void PastePresentationStoreSlide(int index)
        {
            logger.Log("Trying to paste slide " + index);
            var presentationDir = GetPresentationDir();
            var presentations = Directory.GetDirectories(presentationDir);
            int currentIndex = 0;
            foreach (var presentationUUID in presentations)
            {
                logger.Log("Searching in presentation: " + presentationUUID);
                var temporaryPresentation = GetStoredPresentationByFolder(presentationUUID);
                int slideIndex = 0;
                foreach (var slide in temporaryPresentation.Slides)
                {
                    if (currentIndex == index)
                    {
                        temporaryPresentation.Slides[slideIndex + 1].Copy(); //brillant design decision to start counting at 1.
                        try
                        {
                            var currentlySelectedSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex;
                            Globals.ThisAddIn.Application.ActiveWindow.Presentation.Slides.Paste(currentlySelectedSlide + 1);
                        }
                        catch(Exception e)
                        {
                            logger.Log("Could not determine current slide: " + e);
                            Globals.ThisAddIn.Application.ActiveWindow.Presentation.Slides.Paste(); //use default index as fallback
                        }
                        return;
                    }
                    else
                    {
                        currentIndex++;
                    }
                    slideIndex++;
                }
                temporaryPresentation.Close();
            }
            throw new ArgumentOutOfRangeException("No such presentation slide index.");
        }

        internal Bitmap GetPresentationStoreSlideThumbByIndex(int index)
        {
            var presentationDir = GetPresentationDir();
            var presentations = Directory.GetDirectories(presentationDir);
            int currentIndex = 0;
            foreach (var presentationUUID in presentations)
            {
                var temporaryPresentation = GetStoredPresentationByFolder(presentationUUID);
                int slideIndex = 0;
                foreach(var slide in temporaryPresentation.Slides)
                {
                    if(currentIndex == index)
                    {
                        return GetPresentationStoreSlideThumbByFolderAndIndex(presentationUUID, slideIndex);
                    }
                    else
                    {
                        currentIndex++;
                    }
                    slideIndex++;
                }
                temporaryPresentation.Close();
            }
            throw new ArgumentOutOfRangeException("No such presentation slide index.");
        }

        private Bitmap GetPresentationStoreSlideThumbByFolderAndIndex(string presentationFolder, int slideIndex)
        {
            var thumbs = Directory.GetFiles(presentationFolder, "*.png");
            var targetThumbFile = thumbs[slideIndex];
            using(var image = Image.FromFile(targetThumbFile, true))
            {
                return new Bitmap(image);
            }
        }

        internal int GetPresentationStoreSlideCount()
        {
            var result = 0;
            var presentationDir = GetPresentationDir();
            var presentations = Directory.GetDirectories(presentationDir);
            foreach(var presentationUUID in presentations)
            {
                var temporaryPresentation = GetStoredPresentationByFolder(presentationUUID);
                result += temporaryPresentation.Slides.Count;
                temporaryPresentation.Close();
            }
            return result;
        }

        //do not forget to close the obtained presentation
        internal PowerPoint.Presentation GetStoredPresentationByFolder(string uuid)
        {
            var presentationFile = uuid + Path.DirectorySeparatorChar + STORED_PRESENTATION_FILENAME;
            return Globals.ThisAddIn.Application.Presentations.Open(presentationFile, Core.MsoTriState.msoTrue, Core.MsoTriState.msoTrue, Core.MsoTriState.msoFalse);
        }

        //checks whether the two specified shape types belong to the same group in SupportedShapeTypes
        private bool IsPeerShapeType(Core.MsoShapeType shapeType1, Core.MsoShapeType shapeType2)
        {
            foreach (List<Core.MsoShapeType> shapeTypeGroup in SupportedShapeTypes.All.Values)
            {
                if (shapeTypeGroup.Contains(shapeType1) && shapeTypeGroup.Contains(shapeType2))
                {
                    return true;
                }
            }
            return false;
        }

        public bool SaveShapeFromClipBoard(Shape shape)
        {
            var temporaryPresentation = Globals.ThisAddIn.Application.Presentations.Add(Core.MsoTriState.msoFalse);
            var targetSlide = temporaryPresentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            var newUuid = Guid.NewGuid().ToString();
            var fileName = GetFileName(newUuid);
            var persistenceFile = GetPersistenceFile(fileName);
            try
            {
                targetSlide.Shapes.Paste();
            }
            catch (Exception e)
            {
                logger.Log("Clipboard content could not be pasted. " + e.Message);
                temporaryPresentation.Close();
                return false;
            }
            logger.Log("Saving shape.");
            var cachedShapes = CachedShapes; // ensure this is already loaded before saving!
            logger.Log("Saving shape of type " + targetSlide.Shapes[1]);
            temporaryPresentation.SaveAs(persistenceFile, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Core.MsoTriState.msoFalse);
            temporaryPresentation.Close();
            //we reload this since the shape's type seems not to be known otherwise...
            var shapeToSave = new ShapeFavorite(persistenceFile, this.GetShapesFromFile(persistenceFile).First()); // there is only one shape in this file
            if (shape == null)
            {
                this.GetThumbnail(shapeToSave); //create thumbnail
            }
            else
            {
                this.CreateThumbnail(GetThumbnailPath(shapeToSave), shape);
            }
            cachedShapes.Add(shapeToSave);
            InformCacheListenersOnItemAdded(shapeToSave);
            return true;
        }

        public void SaveShape(Shape shape)
        {
            shape.Copy();
            this.SaveShapeFromClipBoard(shape);
        }


        public List<ShapeFavorite> GetShapes()
        {
            return new List<ShapeFavorite>(CachedShapes);
        }


        internal void LoadShapes()
        {
            var persistenceDir = GetPersistenceDir();
            logger.Log("Loading shapes from persistence directory: " + persistenceDir);
            string[] filePaths = Directory.GetFiles(@persistenceDir, "*" + PERSISTENCE_EXTENSION,
                                         System.IO.SearchOption.TopDirectoryOnly).OrderBy(path => new FileInfo(path).CreationTime).ToArray();
            var loadedShapes = new List<ShapeFavorite>();
            foreach (string file in filePaths)
            {
                logger.Log("Reading file " + file);
                List<Shape> shapesFromFile = this.GetShapesFromFile(file);
                foreach (Shape shape in shapesFromFile)
                {
                    loadedShapes.Add(new ShapeFavorite(file, shape));
                    logger.Log("Loaded shape from of type " + shape.Type + " from file " + file);
                }
            }
            var structure = LoadStructure();
            var orderedShapes = OrderShapes(loadedShapes, structure);
            CachedShapes = orderedShapes;
        }

        private List<ShapeFavorite> OrderShapes(List<ShapeFavorite> shapes, List<string> structure)
        {
            var allShapes = new List<ShapeFavorite>(shapes);
            var result = new List<ShapeFavorite>();
            foreach (string structureEntry in structure)
            {
                foreach (ShapeFavorite shape in allShapes)
                {
                    if (shape.FilePath.Contains(structureEntry))
                    {
                        result.Add(shape);
                        break;
                    }
                }
            }
            allShapes.RemoveAll(favorite => result.Contains(favorite));
            logger.Log("Found all structure entries except: " + allShapes.Count);
            result.AddRange(allShapes);
            logger.Log("Ordered " + shapes.Count + " shapes. Result has size: " + result.Count);
            return result;
        }

        private List<string> LoadStructure()
        {
            if (!File.Exists(STRUCTURE_PERSISTANCE_FILE))
            {
                logger.Log("Structure file does not exist.");
                return new List<string>();
            }
            var structureString = File.ReadAllText(STRUCTURE_PERSISTANCE_FILE);
            return JsonConvert.DeserializeObject<List<string>>(structureString);
        }

        private List<Shape> GetShapesFromFile(string file)
        {
            var temporaryPresentation = Globals.ThisAddIn.Application.Presentations.Open(file, Core.MsoTriState.msoTrue, Core.MsoTriState.msoTrue, Core.MsoTriState.msoFalse);
            List<Shape> result = new List<Shape>();
            foreach (Shape shape in temporaryPresentation.Slides[1].Shapes.Cast<Shape>().ToList())
            {
                result.Add(shape);
            }
            temporaryPresentation.Close();
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
            logger.Log("Using file path: " + filePath);
            return filePath;
        }

        internal static string GetPresentationDir()
        {
            var persistenceDir = GetPersistenceDir();
            var presentationDir = persistenceDir + Path.DirectorySeparatorChar + PRESENTATION_DIR;
            Directory.CreateDirectory(presentationDir);
            return presentationDir;
        }

        internal static string GetPersistenceDir()
        {
            var homePath = Environment.ExpandEnvironmentVariables("%HOMEDRIVE%%HOMEPATH%");
            var separator = Path.DirectorySeparatorChar;
            var persistenceDir = homePath + separator + PERSISTENCE_DIR;
            Directory.CreateDirectory(persistenceDir);
            return persistenceDir;
        }

        internal void RegisterCacheListener(CacheListener listener)
        {
            logger.Log("Adding cache listener: " + listener);
            this.cacheListeners.Add(listener);
        }

        internal interface CacheListener
        {
            void OnCacheRenewed();

            void OnItemAdded(ShapeFavorite addedItem);

            void OnItemRemoved(ShapeFavorite removedItem);
        }

        private class StructurePersistanceListener : CacheListener
        {
            private readonly ShapePersistence parent;
            private readonly string structureFile;
            private static readonly DebugLogger logger = DebugLogger.GetLogger(typeof(StructurePersistanceListener).Name);

            internal StructurePersistanceListener(ShapePersistence parent, string structureFile)
            {
                this.parent = parent;
                this.structureFile = structureFile;
            }

            public void OnCacheRenewed()
            {
                WriteStructure();
            }

            public void OnItemAdded(ShapeFavorite addedItem)
            {
                WriteStructure();
            }

            public void OnItemRemoved(ShapeFavorite removedItem)
            {
                WriteStructure();
            }

            private void WriteStructure()
            {
                logger.Log("Writing structure file.");
                File.WriteAllText(structureFile, GetStructureJson());
            }

            private string GetStructureJson()
            {
                var shapesList = new List<string>();
                foreach (ShapeFavorite fav in parent.CachedShapes)
                {
                    shapesList.Add(Path.GetFileName(fav.FilePath));
                }
                return JsonConvert.SerializeObject(shapesList);
            }
        }
    }
}
