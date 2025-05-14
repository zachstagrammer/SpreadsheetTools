using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelToolkit
{
    public interface IExcelImporter
    {
        List<T> Import<T>(string filePath, string columnAHeaderName) where T : class, new();
        List<T> Import<T>(string filePath, int headerRowIndex) where T : class, new();
    }
}
