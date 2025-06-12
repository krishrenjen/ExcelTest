using ClosedXML.Excel;

namespace ExcelTest.Services
{
    public class ColumnSchema
    {
        public string Name { get; set; }
        public bool Required { get; set; }
        public int? MaxLength { get; set; }
        public int? MinValue { get; set; }
        public int? MaxValue { get; set; }
    }

    public class ExcelValidator
    {
        private readonly List<ColumnSchema> _schema;

        public ExcelValidator(List<ColumnSchema> schema)
        {
            _schema = schema;
        }

        public (List<Dictionary<string, object>> validRows, List<string> errors) ValidateWorkbook(XLWorkbook workbook)
        {
            var errors = new List<string>();
            var validRows = new List<Dictionary<string, object>>();

            if (workbook.Worksheets.Count == 0)
            {
                errors.Add("The workbook does not contain any worksheets.");
                return (validRows, errors);
            }

            var worksheet = workbook.Worksheet(1);
            if (worksheet.LastRowUsed() == null)
            {
                errors.Add("The worksheet does not contain any data.");
                return (validRows, errors);
            }

            var headerRow = worksheet.Row(1);
            var headerCells = headerRow.CellsUsed().Select(c => c.GetString()).ToList();

            var columnIndices = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headerCells.Count; i++)
            {
                var columnName = headerCells[i];
                if (_schema.Any(s => s.Name.Equals(columnName, StringComparison.OrdinalIgnoreCase)))
                {
                    columnIndices[columnName] = i + 1; // Excel is 1-indexed
                }
            }

            foreach (var column in _schema)
            {
                if (column.Required && !columnIndices.Keys.Any(k => k.Equals(column.Name, StringComparison.OrdinalIgnoreCase)))
                {
                    errors.Add($"Column '{column.Name}' is missing in the header row.");
                }
            }

            if (errors.Count > 0)
            {
                return (validRows, errors);
            }

            var dataRows = worksheet.RowsUsed().Skip(1).ToList();
            if (dataRows.Count == 0)
            {
                errors.Add("The worksheet does not contain any data rows.");
                return (validRows, errors);
            }

            foreach (var row in dataRows)
            {
                var rowNumber = row.RowNumber();
                var rowData = new Dictionary<string, object>();
                bool isValidRow = true;

                foreach (var colSchema in _schema)
                {
                    if (!columnIndices.TryGetValue(colSchema.Name, out int colNumber))
                    {
                        if (colSchema.Required)
                        {
                            errors.Add($"Column '{colSchema.Name}' is missing in the header row.");
                            isValidRow = false;

                            return (validRows, errors);
                        }
                        continue;
                    }

                    var cell = row.Cell(colNumber);
                    var cellValue = cell.GetValue<string>()?.Trim();

                    if (colSchema.Required && string.IsNullOrWhiteSpace(cellValue))
                    {
                        errors.Add($"Row {rowNumber}: '{colSchema.Name}' is required.");
                        isValidRow = false;
                        continue;
                    }

                    if (colSchema.MaxLength.HasValue && !string.IsNullOrWhiteSpace(cellValue) &&
                        cellValue.Length > colSchema.MaxLength.Value)
                    {
                        errors.Add($"Row {rowNumber}: '{colSchema.Name}' exceeds max length of {colSchema.MaxLength}.");
                        isValidRow = false;
                        continue;
                    }

                    if ((colSchema.MinValue.HasValue || colSchema.MaxValue.HasValue) && !string.IsNullOrWhiteSpace(cellValue))
                    {
                        if (int.TryParse(cellValue, out int intValue))
                        {
                            if (colSchema.MinValue.HasValue && intValue < colSchema.MinValue.Value)
                            {
                                errors.Add($"Row {rowNumber}: '{colSchema.Name}' must be at least {colSchema.MinValue}.");
                                isValidRow = false;
                                continue;
                            }
                            if (colSchema.MaxValue.HasValue && intValue > colSchema.MaxValue.Value)
                            {
                                errors.Add($"Row {rowNumber}: '{colSchema.Name}' must be at most {colSchema.MaxValue}.");
                                isValidRow = false;
                                continue;
                            }

                            rowData[colSchema.Name] = intValue;
                        }
                        else
                        {
                            errors.Add($"Row {rowNumber}: '{colSchema.Name}' must be a valid integer.");
                            isValidRow = false;
                            continue;
                        }
                    }
                    else if (!rowData.ContainsKey(colSchema.Name))
                    {
                        rowData[colSchema.Name] = cellValue ?? "";
                    }
                }

                if (isValidRow) {
                    validRows.Add(rowData);
                }
                else if (!isValidRow || errors.Count > 0)
                {
                    return (validRows, errors);
                }

            }

            return (validRows, errors);
        }
    }
}
