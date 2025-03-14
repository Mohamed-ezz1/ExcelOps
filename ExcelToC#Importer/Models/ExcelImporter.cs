//using OfficeOpenXml;

//namespace ExcelImporterMethod
//{
//    public static class ExcelImporter
//    {
//        public static List<T> ReadExcelData<T>(string filePath) where T : new()
//        {
//            List<T> dataList = new List<T>();

//            using (var package = new ExcelPackage(new FileInfo(filePath)))
//            {
//                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
//                var worksheet = package.Workbook.Worksheets[0];
//                int rowCount = worksheet.Dimension.Rows;
//                int colCount = worksheet.Dimension.Columns;

//                var headers = new List<string>();
//                for (int col = 1; col <= colCount; col++)
//                {
//                    headers.Add(worksheet.Cells[1, col].Text);
//                }

//                var properties = typeof(T).GetProperties();

//                for (int row = 2; row <= rowCount; row++)
//                {
//                    T data = new T();

//                    for (int col = 1; col <= colCount; col++)
//                    {
//                        string header = headers[col - 1];
//                        string cellValue = worksheet.Cells[row, col].Text;

//                        var property = properties.FirstOrDefault(p => p.Name.Equals(header, StringComparison.OrdinalIgnoreCase));

//                        if (property is not null && !string.IsNullOrEmpty(cellValue))
//                        {
//                            if (property.PropertyType == typeof(int))
//                            {
//                                property.SetValue(data, int.Parse(cellValue));
//                            }
//                            else
//                            {
//                                property.SetValue(data, cellValue);
//                            }
//                        }
//                    }
//                    dataList.Add(data);
//                }
//            }

//            return dataList;
//        }
//    }
//}
