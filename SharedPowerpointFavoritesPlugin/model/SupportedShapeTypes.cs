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

        public static ReadOnlyCollection<MsoShapeType> Shapes
        {
            get;
        }

        public static ReadOnlyCollection<MsoShapeType> Charts
        {
            get;
        }

        public static ReadOnlyCollection<MsoShapeType> Tables
        {
            get;
        }

        public static ReadOnlyCollection<MsoShapeType> Pictures
        {
            get;
        }

        public static ReadOnlyCollection<MsoShapeType> Others
        {
            get;
        }
        
        static SupportedShapeTypes()
        {
            Shapes = GetList(MsoShapeType.msoAutoShape, MsoShapeType.msoGroup, MsoShapeType.msoFreeform, MsoShapeType.msoTextBox).AsReadOnly();
            Charts = GetList(MsoShapeType.msoChart).AsReadOnly();
            Tables = GetList(MsoShapeType.msoTable).AsReadOnly();
            Pictures = GetList(MsoShapeType.msoPicture).AsReadOnly();

            var _all = new Dictionary<string, List<MsoShapeType>>();
            _all.Add("Shapes", Shapes.ToList());
            _all.Add("Charts", Charts.ToList());
            _all.Add("Tables", Tables.ToList());
            _all.Add("Pictures", Pictures.ToList());
            Others = GetRemainingShapeTypes(_all.Values.SelectMany(i => i)).AsReadOnly();
            _all.Add("Others", Others.ToList());

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
