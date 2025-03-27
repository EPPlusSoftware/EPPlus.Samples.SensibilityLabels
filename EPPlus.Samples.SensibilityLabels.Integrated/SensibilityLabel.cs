using OfficeOpenXml.Interfaces.SensitivityLabels;
using System.Diagnostics;

namespace SensibilityLabelHandler
{
    [DebuggerDisplay("Name: {Name}")]
    public class SensibilityLabel : IExcelSensibilityLabel, IExcelSensibilityLabelUpdate
    {
        public string Id { get; set; }

        public string Name { get; set; }

        public string Tooltip { get; set; }
        
        public string Description { get; set; }

        public string Color { get; set; }

        public bool Enabled { get; set; }

        public bool Removed { get; set; }

        public string SiteId { get; set; }

        public eMethod Method { get; set; }

        public eContentBits ContentBits { get; set; }

        public IExcelSensibilityLabel Parent { get; set; }
        
        public void Update(string name, string tooltip, string description, string color, IExcelSensibilityLabel parent)
        {
            Name = name;
            Tooltip = tooltip;
            Description = description;
            Color = color;
            Parent = parent;
        }
    }
}
