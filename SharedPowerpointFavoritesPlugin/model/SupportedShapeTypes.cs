using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharedPowerpointFavoritesPlugin.model
{
    static class SupportedShapeTypes
    {

        private static readonly DebugLogger logger = DebugLogger.GetLogger(typeof(SupportedShapeTypes).Name);

        public static IDictionary<string, List<MsoShapeType>> All
        {
            get;
        }

        public static List<MsoShapeType> Others
        {
            get;
        }
        
        static SupportedShapeTypes()
        {
            var _all = new Dictionary<string, List<MsoShapeType>>();
            _all.Add("Shapes", GetList(MsoShapeType.msoAutoShape));
            _all.Add("Charts", GetList(MsoShapeType.msoChart));
            _all.Add("Tables", GetList(MsoShapeType.msoTable));
            _all.Add("Pictures", GetList(MsoShapeType.msoPicture));
            _all.Add("Groups", GetList(MsoShapeType.msoGroup));
            var otherShapeTypes = GetRemainingShapeTypes(_all.Values.SelectMany(i => i));
            _all.Add("Others", otherShapeTypes);
            Others = otherShapeTypes;
            All = new ReadOnlyDictionary<string, List<MsoShapeType>>(_all);
        }

        private static List<MsoShapeType> GetList(params MsoShapeType[] shapeTypes)
        {
            var list = new List<MsoShapeType>();
            foreach(MsoShapeType type in shapeTypes)
            {
                list.Add(type);
            }
            return list;
        }

        private static List<MsoShapeType> GetRemainingShapeTypes(IEnumerable<MsoShapeType> notToInclude)
        {
            var otherShapeTypes = new List<MsoShapeType>(Enum.GetValues(typeof(MsoShapeType)).Cast<MsoShapeType>());
            otherShapeTypes.RemoveAll(item => notToInclude.Contains(item));
            foreach(MsoShapeType type in otherShapeTypes)
            {
                logger.Log("Adding Others shapeType: " + type);
            }
            return otherShapeTypes;
        }
    }
}
