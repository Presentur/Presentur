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
using System.Text;
using System.Threading.Tasks;

namespace SharedPowerpointFavoritesPlugin
{
    class DebugLogger
    {
        internal static bool DEBUG_LOGGING_ENABLED = true;
        private readonly string tag;
        private readonly string logFile;

        private DebugLogger(string tag)
        {
            this.tag = tag;
            this.logFile = Path.Combine(ShapePersistence.GetPresenturDir(), "presentur.log");
        }
        public void Log(string message)
        {
            if(!DEBUG_LOGGING_ENABLED)
            {
                return;
            }
            var formattedLogMessage = DateTime.Now.ToString() + " | " + tag + ": " + message;
            System.Diagnostics.Debug.WriteLine(formattedLogMessage);
            File.AppendAllText(this.logFile, formattedLogMessage + Environment.NewLine);
        }

        public static DebugLogger GetLogger(string tag)
        {
            return new DebugLogger(tag);
        }
    }
}
