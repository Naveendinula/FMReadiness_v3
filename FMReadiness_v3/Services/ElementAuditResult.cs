using System.Collections.Generic;

namespace FMReadiness_v3.Services
{
    public class ElementAuditResult
    {
        public int ElementId { get; set; }
        public string Category { get; set; } = string.Empty;
        public string Family { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public int MissingCount { get; set; }
        public double ReadinessScore { get; set; }
        public string MissingParams { get; set; } = string.Empty;

        // Group-level scores (group name -> score 0..1)
        public Dictionary<string, double> GroupScores { get; set; } = new();

        // Missing fields annotated with group: "[Identity] Asset Tag, [Location] Room"
        public List<MissingFieldInfo> MissingFields { get; set; } = new();
    }

    public class MissingFieldInfo
    {
        public string Group { get; set; } = string.Empty;
        public string FieldLabel { get; set; } = string.Empty;
        public string? Reason { get; set; }
    }
}

