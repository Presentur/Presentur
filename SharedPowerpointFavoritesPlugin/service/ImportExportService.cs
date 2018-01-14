using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;

namespace SharedPowerpointFavoritesPlugin
{
    class ImportExportService
    {
        public static ImportExportService INSTANCE = new ImportExportService();
        private ShapePersistence shapePersistance = ShapePersistence.INSTANCE;
        private static readonly DebugLogger logger = DebugLogger.GetLogger(typeof(ImportExportService).Name);

        private ImportExportService()
        {
            //singleton
        }

        public bool ImportFromFile(string filePath)
        {
            try
            {
                var persistenceDir = this.shapePersistance.GetPersistenceDir();
                Directory.Delete(persistenceDir, true);
                Directory.CreateDirectory(persistenceDir);
                ZipFile.ExtractToDirectory(filePath, persistenceDir);
                this.shapePersistance.LoadShapes();
            }
            catch(Exception e)
            {
                logger.Log("Exception while importing from file: " + e.Message);
                return false;
            }
            logger.Log("Import successful.");
            return true;
        }

        public bool ExportToFile(string filePath)
        {
            var dataPath = this.shapePersistance.GetPersistenceDir();
            try
            {
                ZipFile.CreateFromDirectory(dataPath, filePath, CompressionLevel.Fastest, false);
            }
            catch(Exception e)
            {
                logger.Log("Exception while exporting: " + e.Message);
                return false;
            }
            logger.Log("Export successful.");
            return true;
        }
    }
}
