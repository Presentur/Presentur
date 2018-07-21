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
using System.Windows.Forms;

namespace SharedPowerpointFavoritesPlugin.util
{
    static class DialogUtil
    {

        private static readonly DebugLogger logger = DebugLogger.GetLogger(typeof(DialogUtil).Name);

        public const string POWERPOINT_PRESENTATION_FILTER = "Powerpoint Presentation (*.pptx)|*.pptx";

        public static bool AskForConfirmation(string message)
        {
            return MessageBox.Show(message,
                                     "Confirm",
                                     MessageBoxButtons.YesNo) == DialogResult.Yes;
        }

        public static string GetFilePathViaDialog(bool isSaveAction, string filter)
        {
            FileDialog openFileDialog = isSaveAction ? (new SaveFileDialog() as FileDialog) : (new OpenFileDialog() as FileDialog);
            openFileDialog.InitialDirectory = Environment.ExpandEnvironmentVariables("%HOMEDRIVE%%HOMEPATH%");
            openFileDialog.Filter = filter;
            openFileDialog.RestoreDirectory = true;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog.FileName;
            }
            else
            {
                logger.Log("No file chosen.");
                return null;
            }
        }

        public static string GetFilePathViaDialog(bool isSaveAction)
        {
            return GetFilePathViaDialog(isSaveAction, "SharedPowerpointFavorites (*.zip)|*.zip");
        }
    }
}
