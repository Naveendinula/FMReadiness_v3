using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace FMReadiness_v3.Services
{
    public class CollectorService
    {
        private readonly Document _doc;

        public CollectorService(Document doc)
        {
            _doc = doc;
        }

        public List<Element> GetAllFmElements()
        {
            var categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_DuctTerminal,
                BuiltInCategory.OST_DuctAccessory,
                BuiltInCategory.OST_PipeAccessory
            };

            var filter = new ElementMulticategoryFilter(categories);

            return new FilteredElementCollector(_doc)
                .WherePasses(filter)
                .WhereElementIsNotElementType()
                .ToElements()
                .ToList();
        }
    }
}

