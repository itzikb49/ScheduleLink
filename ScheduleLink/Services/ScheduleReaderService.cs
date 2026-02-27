using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using ScheduleLink.Helpers;
using ScheduleLink.Models;

namespace ScheduleLink.Services
{
    public static class ScheduleReaderService
    {
        public static List<ViewSchedule> GetAllSchedules(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(vs => !vs.IsTemplate
                          && !vs.IsTitleblockRevisionSchedule
                          && vs.Definition != null)
                .OrderBy(vs => vs.Name)
                .ToList();
        }

        public static ScheduleExportData ExtractScheduleData(Document doc, ViewSchedule schedule)
        {
            ScheduleDefinition definition = schedule.Definition;
            var exportData = new ScheduleExportData
            {
                ScheduleName = schedule.Name,
                ScheduleId = schedule.Id.GetIdValueLong()
            };

            // --- 1. Extract visible field definitions ---
            int fieldCount = definition.GetFieldCount();
            var activeFields = new List<ScheduleField>();

            for (int i = 0; i < fieldCount; i++)
            {
                ScheduleField field = definition.GetField(i);
                if (field.IsHidden)
                    continue;

                var colInfo = new ScheduleColumnInfo
                {
                    FieldIndex = i,
                    Name = field.GetName(),
                    HeaderText = field.ColumnHeading,
                    IsHidden = false,
                };

                ScheduleFieldType fieldType = field.FieldType;

                colInfo.IsCalculated = (fieldType == ScheduleFieldType.Formula
                                     || fieldType == ScheduleFieldType.Count);

                if (fieldType == ScheduleFieldType.ElementType)
                {
                    colInfo.IsTypeParam = true;
                    colInfo.IsReadOnly = false;
                }
                else if (fieldType == ScheduleFieldType.Instance)
                {
                    colInfo.IsTypeParam = false;
                    colInfo.IsReadOnly = false;
                }
                else
                {
                    colInfo.IsReadOnly = true;
                    colInfo.IsCalculated = true;
                }

                exportData.Columns.Add(colInfo);
                activeFields.Add(field);
            }

            // --- 2. Get ALL elements in schedule ---
            List<ElementId> allElementIds = GetScheduleElementIds(doc, schedule);

            // --- 3. Read data per element (not per table row) ---
            // This approach avoids issues with grouped/sorted schedules
            // where table rows don't match element order
            foreach (ElementId eid in allElementIds)
            {
                Element elem = doc.GetElement(eid);
                if (elem == null)
                    continue;

                var rowData = new ScheduleRowData();
                rowData.ElementId = eid.GetIdValueLong();

                for (int c = 0; c < exportData.Columns.Count; c++)
                {
                    var colInfo = exportData.Columns[c];

                    // Find parameter by column heading first, then by field name
                    Parameter param = FindParameter(elem, colInfo.HeaderText);
                    if (param == null && colInfo.Name != colInfo.HeaderText)
                        param = FindParameter(elem, colInfo.Name);

                    if (param != null && param.HasValue)
                    {
                        rowData.Values.Add(GetParameterDisplayValue(param));
                    }
                    else
                    {
                        rowData.Values.Add(string.Empty);
                    }
                }

                exportData.Rows.Add(rowData);
            }

            // --- 4. Verify read-only + storage type using first element ---
            if (allElementIds.Count > 0)
            {
                Element firstElem = doc.GetElement(allElementIds[0]);
                if (firstElem != null)
                {
                    for (int c = 0; c < exportData.Columns.Count; c++)
                    {
                        if (exportData.Columns[c].IsCalculated)
                            continue;

                        var col = exportData.Columns[c];
                        Parameter param = FindParameter(firstElem, col.HeaderText)
                                       ?? FindParameter(firstElem, col.Name);

                        if (param != null)
                        {
                            col.IsReadOnly = param.IsReadOnly;
                            col.StorageType = param.StorageType.ToString();
                        }
                        else
                        {
                            col.IsReadOnly = true;
                        }
                    }
                }
            }

            return exportData;
        }

        private static List<ElementId> GetScheduleElementIds(Document doc, ViewSchedule schedule)
        {
            try
            {
                var collector = new FilteredElementCollector(doc, schedule.Id);
                return collector.ToElementIds().ToList();
            }
            catch
            {
                return new List<ElementId>();
            }
        }

        /// <summary>
        /// Gets display value from parameter (as shown in Revit).
        /// </summary>
        private static string GetParameterDisplayValue(Parameter param)
        {
            if (param == null || !param.HasValue)
                return string.Empty;

            // AsValueString gives the display value (with units, formatted)
            string vs = param.AsValueString();
            if (!string.IsNullOrEmpty(vs))
                return vs;

            switch (param.StorageType)
            {
                case StorageType.String:
                    return param.AsString() ?? string.Empty;
                case StorageType.Integer:
                    return param.AsInteger().ToString();
                case StorageType.Double:
                    return param.AsDouble().ToString(System.Globalization.CultureInfo.CurrentCulture);
                case StorageType.ElementId:
                    return param.AsElementId().GetIdValueLong().ToString();
                default:
                    return string.Empty;
            }
        }

        public static Parameter FindParameter(Element elem, string paramName)
        {
            if (elem == null || string.IsNullOrEmpty(paramName))
                return null;

            Parameter p = elem.LookupParameter(paramName);
            if (p != null) return p;

            foreach (Parameter param in elem.Parameters)
            {
                if (param.Definition != null &&
                    param.Definition.Name.Equals(paramName, StringComparison.OrdinalIgnoreCase))
                    return param;
            }

            ElementId typeId = elem.GetTypeId();
            if (typeId != null && typeId != ElementId.InvalidElementId)
            {
                Element typeElem = elem.Document.GetElement(typeId);
                if (typeElem != null)
                {
                    p = typeElem.LookupParameter(paramName);
                    if (p != null) return p;

                    foreach (Parameter param in typeElem.Parameters)
                    {
                        if (param.Definition != null &&
                            param.Definition.Name.Equals(paramName, StringComparison.OrdinalIgnoreCase))
                            return param;
                    }
                }
            }

            return null;
        }
    }
}
