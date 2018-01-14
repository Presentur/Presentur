using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharedPowerpointFavoritesPlugin
{
    class DebugLogger
    {
        internal static bool DEBUG_LOGGING_ENABLED = true;
        private readonly string tag;

        private DebugLogger(string tag)
        {
            this.tag = tag;
        }
        public void Log(string message)
        {
            if(!DEBUG_LOGGING_ENABLED)
            {
                return;
            }
            System.Diagnostics.Debug.WriteLine(DateTime.Now.ToString() + " | " + tag + ": " + message);
        }

        public static DebugLogger GetLogger(string tag)
        {
            return new DebugLogger(tag);
        }
    }
}
