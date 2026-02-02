using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;

namespace FMReadiness_v3.Services
{
    public class CollectorService
    {
        private readonly Document _doc;

        public CollectorService(Document doc)
        {
            _doc = doc;
        }

        /// <summary>
        /// Gets all FM-relevant elements (MEP equipment, accessories, etc.).
        /// </summary>
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

        /// <summary>
        /// Gets all FM elements plus extended COBie categories (electrical, plumbing, etc.).
        /// </summary>
        public List<Element> GetAllCobieComponents()
        {
            var categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_DuctTerminal,
                BuiltInCategory.OST_DuctAccessory,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_ElectricalEquipment,
                BuiltInCategory.OST_ElectricalFixtures,
                BuiltInCategory.OST_LightingFixtures,
                BuiltInCategory.OST_FireAlarmDevices,
                BuiltInCategory.OST_Furniture
            };

            var filter = new ElementMulticategoryFilter(categories);

            return new FilteredElementCollector(_doc)
                .WherePasses(filter)
                .WhereElementIsNotElementType()
                .ToElements()
                .ToList();
        }

        /// <summary>
        /// Gets all Rooms in the document for COBie Space table.
        /// </summary>
        public List<Room> GetAllRooms()
        {
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(SpatialElement))
                .OfType<Room>()
                .Where(r => r.Area > 0) // Only placed rooms
                .ToList();
        }

        /// <summary>
        /// Gets all MEP Spaces in the document for COBie Space table.
        /// </summary>
        public List<Space> GetAllSpaces()
        {
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(SpatialElement))
                .OfType<Space>()
                .Where(s => s.Area > 0) // Only placed spaces
                .ToList();
        }

        /// <summary>
        /// Gets all Levels in the document for COBie Floor table.
        /// </summary>
        public List<Level> GetAllLevels()
        {
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .OrderBy(l => l.Elevation)
                .ToList();
        }

        /// <summary>
        /// Gets combined Rooms and Spaces for COBie Space coverage.
        /// </summary>
        public List<Element> GetAllSpatialElements()
        {
            var elements = new List<Element>();
            elements.AddRange(GetAllRooms());
            elements.AddRange(GetAllSpaces());
            return elements;
        }

        /// <summary>
        /// Gets elements by specific categories.
        /// </summary>
        public List<Element> GetElementsByCategories(IEnumerable<BuiltInCategory> categories)
        {
            var filter = new ElementMulticategoryFilter(categories.ToList());

            return new FilteredElementCollector(_doc)
                .WherePasses(filter)
                .WhereElementIsNotElementType()
                .ToElements()
                .ToList();
        }

        /// <summary>
        /// Gets all unique element types used by the collected components.
        /// </summary>
        public List<ElementType> GetComponentTypes(IEnumerable<Element> components)
        {
            var typeIds = components
                .Select(e => e.GetTypeId())
                .Where(id => id != ElementId.InvalidElementId)
                .Distinct()
                .ToList();

            return typeIds
                .Select(id => _doc.GetElement(id) as ElementType)
                .Where(t => t != null)
                .ToList()!;
        }

        /// <summary>
        /// Gets COBie collection summary for the document.
        /// </summary>
        public CobieCollectionSummary GetCobieSummary()
        {
            return new CobieCollectionSummary
            {
                ComponentCount = GetAllCobieComponents().Count,
                RoomCount = GetAllRooms().Count,
                SpaceCount = GetAllSpaces().Count,
                LevelCount = GetAllLevels().Count,
                TypeCount = GetComponentTypes(GetAllCobieComponents()).Count
            };
        }
    }

    public class CobieCollectionSummary
    {
        public int ComponentCount { get; set; }
        public int RoomCount { get; set; }
        public int SpaceCount { get; set; }
        public int LevelCount { get; set; }
        public int TypeCount { get; set; }

        public int TotalSpatialElements => RoomCount + SpaceCount;
    }
}

