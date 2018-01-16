using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharedPowerpointFavoritesPlugin.util
{
    static class BuildEnvironment
    {
        public static bool IsAdminBuild()
        {
#if ADMIN
            return true;
#else
            return false;
#endif
        }
    }
}
