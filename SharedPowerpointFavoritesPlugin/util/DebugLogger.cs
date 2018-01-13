using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharedPowerpointFavoritesPlugin
{
    class DebugLogger
    {
        public static void Log(string message)
        {
            System.Diagnostics.Debug.WriteLine(message);
        }
    }
}
