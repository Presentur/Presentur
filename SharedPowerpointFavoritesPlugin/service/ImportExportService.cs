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
                var persistenceDir = ShapePersistence.GetPersistenceDir();
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
            var dataPath = ShapePersistence.GetPersistenceDir();
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
