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

        public static bool AskForConfirmation(string message)
        {
            return MessageBox.Show(message,
                                     "Confirm",
                                     MessageBoxButtons.YesNo) == DialogResult.Yes;
        }

        public static string GetFilePathViaDialog(bool isSaveAction)
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
                logger.Log("No file chosen.");
                return null;
            }
        }
    }
}
