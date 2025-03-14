using ClosedXML.Excel;
using DataHandler;
using System.Reflection;

namespace ExcelImporterMethod
{
    public static class ExcelImporter
    {
        public static List<T> ReadExcelDataUsingPropertyNames<T>(string filePath) where T : new()
        {
            List<T> dataList = new List<T>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RowsUsed();

                var headers = new List<string>();
                var firstRow = rows.First();
                foreach (var cell in firstRow.Cells())
                {
                    headers.Add(cell.Value.ToString());
                }

                var properties = typeof(T).GetProperties();

                foreach (var row in rows.Skip(1))
                {
                    T data = new T();

                    for (int col = 1; col <= headers.Count; col++)
                    {
                        string header = headers[col - 1];
                        string cellValue = row.Cell(col).Value.ToString();

                        var property = properties.FirstOrDefault(p => p.Name.Equals(header, StringComparison.OrdinalIgnoreCase));

                        if (property != null && !string.IsNullOrEmpty(cellValue))
                        {
                            if (property.PropertyType == typeof(int))
                            {
                                property.SetValue(data, int.Parse(cellValue));
                            }
                            else if (property.PropertyType == typeof(double))
                            {
                                property.SetValue(data, double.Parse(cellValue));
                            }
                            else if (property.PropertyType == typeof(DateTime))
                            {
                                property.SetValue(data, DateTime.Parse(cellValue));
                            }
                            else
                            {
                                property.SetValue(data, cellValue);
                            }
                        }
                    }

                    dataList.Add(data);
                }
            }

            return dataList;
        }

        public static List<T> ReadExcelDataUsingAttributes<T>(string filePath) where T : new()
        {
            List<T> dataList = new List<T>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                var rows = worksheet.RowsUsed();

                var headers = new List<string>();
                var firstRow = rows.First();
                foreach (var cell in firstRow.Cells())
                {
                    headers.Add(cell.Value.ToString().Trim());
                }

                var properties = typeof(T).GetProperties();
                var columnMappings = new Dictionary<string, PropertyInfo>(StringComparer.OrdinalIgnoreCase);

                foreach (var property in properties)
                {
                    var attribute = property.GetCustomAttribute<ExcelColumnAttribute>();
                    string columnName = attribute != null ? attribute.ColumnName : property.Name;

                    columnMappings[columnName.Trim()] = property;
                }

                foreach (var row in rows.Skip(1))
                {
                    T data = new T();

                    for (int col = 1; col <= headers.Count; col++)
                    {
                        string header = headers[col - 1];
                        var cell = row.Cell(col);

                        if (cell.IsEmpty() || cell?.Value is null)
                        {
                            continue;
                        }

                        string cellValue = cell.Value.ToString().Trim();

                        if (columnMappings.TryGetValue(header, out var property))
                        {
                            if (!string.IsNullOrEmpty(cellValue))
                            {
                                try
                                {
                                    if (property.PropertyType == typeof(int))
                                    {
                                        property.SetValue(data, int.Parse(cellValue));
                                    }
                                    else if (property.PropertyType == typeof(double))
                                    {
                                        property.SetValue(data, double.Parse(cellValue));
                                    }
                                    else if (property.PropertyType == typeof(DateTime))
                                    {
                                        property.SetValue(data, DateTime.Parse(cellValue));
                                    }
                                    else
                                    {
                                        property.SetValue(data, cellValue);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error converting value '{cellValue}' for property '{property.Name}': {ex.Message}");
                                }
                            }
                        }
                    }

                    dataList.Add(data);
                }
            }

            return dataList;
        }
    }
}


