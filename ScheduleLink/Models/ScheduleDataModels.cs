using System.Collections.Generic;

namespace ScheduleLink.Models
{
    public class ScheduleColumnInfo
    {
        public int FieldIndex { get; set; }
        public string Name { get; set; }
        public string HeaderText { get; set; }
        public bool IsReadOnly { get; set; }
        public bool IsHidden { get; set; }
        public bool IsCalculated { get; set; }
        public string StorageType { get; set; }
        public bool IsTypeParam { get; set; }
    }

    public class ScheduleRowData
    {
        public long ElementId { get; set; }
        public List<string> Values { get; set; } = new List<string>();
    }

    public class ScheduleExportData
    {
        public string ScheduleName { get; set; }
        public long ScheduleId { get; set; }
        public List<ScheduleColumnInfo> Columns { get; set; } = new List<ScheduleColumnInfo>();
        public List<ScheduleRowData> Rows { get; set; } = new List<ScheduleRowData>();
    }

    public class ImportResult
    {
        public int TotalRows { get; set; }
        public int UpdatedParams { get; set; }
        public int SkippedReadOnly { get; set; }
        public int SkippedUnchanged { get; set; }
        public int ParamNotFound { get; set; }
        public int FailedParams { get; set; }
        public int ElementsNotFound { get; set; }
        public List<string> Errors { get; set; } = new List<string>();
    }
}
