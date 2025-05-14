using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ExcelToolkit
{
    public class ExcelImporter : IExcelImporter
    {
        private static bool _encodingRegistered = false;

        /// <summary>
        /// This method processes only the **first worksheet** of the Excel file (.xls or .xlsx).
        /// The worksheet will use the first row as its header row.
        /// Property mapping is based on matching the column headers to either the property names or their <see cref="System.ComponentModel.DisplayNameAttribute"/> values on <typeparamref name="T"/>.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public List<T> Import<T>(string filePath) where T : class, new()
        {
            RegisterEncodingProvider();
            ImposeFileSizeLimit(filePath);

            using (var stream = OpenFileStream(filePath))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var dataTable = GetTable(reader, useHeaderRow: true);
                var columnMap = BuildColumnMap(dataTable.Columns);
                var propertyMap = MapPropertiesToColumns<T>(columnMap);

                return ReadDataRows<T>(dataTable, 1, propertyMap, columnMap).ToList();
            }
        }

        /// <summary>
        /// This method processes only the **first worksheet** of the Excel file (.xls or .xlsx).
        /// The worksheet must contain a clearly defined header row, where one of the cells in column A matches the specified <paramref name="columnAHeaderName"/>.
        /// Property mapping is based on matching the column headers to either the property names or their <see cref="System.ComponentModel.DisplayNameAttribute"/> values on <typeparamref name="T"/>.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <param name="columnAHeaderName"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public List<T> Import<T>(string filePath, string columnAHeaderName) where T : class, new()
        {
            RegisterEncodingProvider();
            ImposeFileSizeLimit(filePath);

            using (var stream = OpenFileStream(filePath))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var dataTable = GetTable(reader);
                int headerRowIndex = FindHeaderRowIndex(dataTable, columnAHeaderName);

                if (headerRowIndex == -1)
                {
                    throw new Exception($"Header row with '{columnAHeaderName}' not found in column A.");
                }

                var columnMap = BuildColumnMap(dataTable.Rows[headerRowIndex]);
                var propertyMap = MapPropertiesToColumns<T>(columnMap);

                return ReadDataRows<T>(dataTable, headerRowIndex + 1, propertyMap, columnMap).ToList();
            }
        }

        /// <summary>
        /// This method processes only the **first worksheet** of the Excel file (.xls or .xlsx).
        /// It uses the specified row index to identify the header row, which is then used to map columns
        /// to properties of <typeparamref name="T"/>.
        /// Property mapping is based on matching the column headers to either the property names or their <see cref="System.ComponentModel.DisplayNameAttribute"/> values on <typeparamref name="T"/>.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <param name="headerRowIndex"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public List<T> Import<T>(string filePath, int headerRowIndex) where T : class, new()
        {
            RegisterEncodingProvider();
            ImposeFileSizeLimit(filePath);

            using (var stream = OpenFileStream(filePath))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var dataTable = GetTable(reader);

                if (headerRowIndex <= 0)
                {
                    throw new Exception($"Header row index '{headerRowIndex}' is invalid");
                }

                var columnMap = BuildColumnMap(dataTable.Rows[headerRowIndex]);
                var propertyMap = MapPropertiesToColumns<T>(columnMap);

                return ReadDataRows<T>(dataTable, headerRowIndex + 1, propertyMap, columnMap).ToList();
            }
        }

        private void RegisterEncodingProvider()
        {
            if (!_encodingRegistered)
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                _encodingRegistered = true;
            }
        }

        private static void ImposeFileSizeLimit(string filePath, long maxSizeInMB = 30)
        {
            var fileInfo = new FileInfo(filePath);
            long fileSizeMB = fileInfo.Length / (1024 * 1024);

            if (fileSizeMB > maxSizeInMB)
            {
                throw new InvalidOperationException($"File size exceeds limit of {maxSizeInMB} MB. Current size: {fileSizeMB} MB");
            }
        }

        private FileStream OpenFileStream(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("Excel file not found.", filePath);
            }

            string extension = Path.GetExtension(filePath)?.ToLowerInvariant();
            if (extension != ".xls" && extension != ".xlsx")
            {
                throw new InvalidOperationException("Unsupported file type. Must be .xls or .xlsx.");
            }

            return File.Open(filePath, FileMode.Open, FileAccess.Read);
        }

        private DataTable GetTable(IExcelDataReader reader, bool useHeaderRow = false)
        {
            var result = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = useHeaderRow }
            });

            return result.Tables[0];
        }

        private int FindHeaderRowIndex(DataTable table, string columnAHeaderName)
        {
            for (int i = 0; i < table.Rows.Count; i++)
            {
                var cell = table.Rows[i][0]?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(cell) && cell.Equals(columnAHeaderName, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        private Dictionary<string, int> BuildColumnMap(DataRow headerRow)
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < headerRow.Table.Columns.Count; i++)
            {
                var columnName = headerRow[i]?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(columnName))
                {
                    map[columnName] = i;
                }
            }

            return map;
        }

        private Dictionary<string, int> BuildColumnMap(DataColumnCollection dataTableColumns)
        {
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < dataTableColumns.Count; i++)
            {
                var columnName = dataTableColumns[i].ColumnName?.Trim();
                if (!string.IsNullOrEmpty(columnName))
                {
                    map[columnName] = i;
                }
            }

            return map;
        }

        private Dictionary<string, PropertyInfo> MapPropertiesToColumns<T>(Dictionary<string, int> columnMap)
        {
            var props = typeof(T).GetProperties();
            var map = new Dictionary<string, PropertyInfo>();

            foreach (var prop in props)
            {
                var display = prop.GetCustomAttribute<DisplayNameAttribute>();
                var columnName = display?.DisplayName ?? prop.Name;

                if (columnMap.ContainsKey(columnName))
                {
                    map[columnName] = prop;
                }
            }

            return map;
        }

        private IEnumerable<T> ReadDataRows<T>(
            DataTable table,
            int startRowIndex,
            Dictionary<string, PropertyInfo> propertyMap,
            Dictionary<string, int> columnMap
        ) where T : new()
        {
            var results = new List<T>();

            for (int rowIndex = startRowIndex; rowIndex < table.Rows.Count; rowIndex++)
            {
                var row = table.Rows[rowIndex];
                if (row.ItemArray.All(cell => string.IsNullOrWhiteSpace(cell?.ToString())))
                {
                    continue;
                }

                var obj = new T();
                foreach (var kvp in propertyMap)
                {
                    var columnIndex = columnMap[kvp.Key];
                    var value = row[columnIndex]?.ToString().Trim();
                    kvp.Value.SetValue(obj, value);
                }

                results.Add(obj);
            }

            return results;
        }
    }
}
