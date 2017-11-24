using CcExcelWriter;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.CSharp;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CoeCall
{
    public class ExcelReader : IDisposable
    {
        private readonly bool _ownStream;
        private readonly Stream _stream;
        private readonly SpreadsheetDocument _document;
        private readonly WorkbookPart _wbPart;
        private readonly DocumentFormat.OpenXml.Spreadsheet.Sheet _sheet;
        private readonly WorksheetPart _wsPart;
        private readonly Cell[] _cells;
        private readonly SharedStringTablePart _stringTable;
        private readonly CSharpCodeProvider _provider;

        public ExcelReader(string fileName) : this(new FileStream(fileName, FileMode.Open, FileAccess.Read))
        {
            _ownStream = true;
        }

        public ExcelReader(Stream stream)
        {
            _provider = new CSharpCodeProvider();
            _stream = stream;
            _document = SpreadsheetDocument.Open(_stream, false);
            _wbPart = _document.WorkbookPart;

            var sheets = _wbPart
                .Workbook
                .Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                .ToList();

            if (sheets.Count != 1) throw new Exception("Multiple sheets not supported.");

            _sheet = sheets.First();

            _wsPart = _wbPart.GetPartById(_sheet.Id) as WorksheetPart;

            if (_wsPart == null) throw new Exception("Excel error: WorksheetPart not found.");

            _cells = _wsPart
                .Worksheet
                .Descendants<Cell>()
                .ToArray();

            _stringTable = _wbPart
                .GetPartsOfType<SharedStringTablePart>()
                .FirstOrDefault();
        }

        public Cell GetCell(string cellReference)
        {
            return _cells
                .Where(w => w.CellReference == cellReference)
                .FirstOrDefault();
        }

        public IEnumerable<Cell> GetCells()
        {
            return _wsPart.Worksheet
                .Descendants<Cell>()
                .ToList();
        }

        public string GetValue(BaseAZ column, uint line)
        {
            var cell = GetCell(column.ToString() + line);

            if (cell?.DataType?.Value == CellValues.SharedString)
            {
                return _stringTable?.SharedStringTable.ElementAt(int.Parse(cell.InnerText)).InnerText;
            }

            return cell?.InnerText;
        }

        public string GetValue(uint column, uint line)
        {
            return GetValue((BaseAZ)column, line);
        }

        public string GetValue(string column, uint line)
        {
            return GetValue(BaseAZ.Parse(column), line);
        }

        public T GetValue<T>(uint column, uint line)
        {
            return GetValue<T>((BaseAZ)column, line);
        }

        public T GetValue<T>(string column, uint line)
        {
            return GetValue<T>(BaseAZ.Parse(column), line);
        }

        public T GetValue<T>(BaseAZ column, uint line)
        {
            var strValue = GetValue(column, line);

            if (string.IsNullOrEmpty(strValue)) return default(T);
            
            if (typeof(T) == typeof(string)) return (T)(object)strValue;


            var enUs = new CultureInfo("en-US");
            var style = NumberStyles.Any;

            var aliasType = GetAlias(typeof(T));

            switch (aliasType)
            {
                case "bool":
                case "bool?":
                    if (int.TryParse(strValue, style, enUs, out var boolValue)) return (T)(object)(boolValue == 1);
                    break;
                case "char":
                case "char?":
                    return (T)(object)strValue?.First();
                case "short":
                case "short?":
                    if (short.TryParse(strValue, style, enUs, out var shortValue)) return (T)(object)shortValue;
                    break;
                case "int":
                case "int?":
                    if (int.TryParse(strValue, style, enUs, out var intValue)) return (T)(object)intValue;
                    break;
                case "long":
                case "long?":
                    if (long.TryParse(strValue, style, enUs, out var longValue)) return (T)(object)longValue;
                    break;
                case "float":
                case "float?":
                    if (float.TryParse(strValue, style, enUs, out var floatValue)) return (T)(object)floatValue;
                    break;
                case "double":
                case "double?":
                    if (double.TryParse(strValue, style, enUs, out var doubleValue)) return (T)(object)doubleValue;
                    break;
                case "decimal":
                case "decimal?":
                    if (decimal.TryParse(strValue, style, enUs, out var decimalValue)) return (T)(object)decimalValue;
                    break;
                case "DateTime":
                case "DateTime?":
                    if (double.TryParse(strValue, style, enUs, out var dateTimeValue)) return (T)(object)DateTime.FromOADate(dateTimeValue);
                    break;
                case "TimeSpan":
                case "TimeSpan?":
                    if (double.TryParse(strValue, style, enUs, out var timeSpanValue))
                    {
                        var dateTime = DateTime.FromOADate(timeSpanValue);
                        return (T)(object)(dateTime - dateTime.Date);
                    }
                    break;
                default:
                    throw new NotSupportedException($"GetValue has no support to {typeof(T)}.");
            }

            return (T)(object)null;
        }

        private string GetAlias(Type type)
        {
            var typeRef = new CodeTypeReference(type);

            var typeName = _provider.GetTypeOutput(typeRef);

            typeName = Regex.Replace(typeName, @"(\w+[.])*(?<name>\w+)", "${name}");
            typeName = Regex.Replace(typeName, @"Nullable[<](?<name>\w+)[>]", "${name}?");

            return typeName;
        }

        public void Dispose()
        {
            _provider?.Dispose();
            if (_ownStream) _stream?.Dispose();
            _document?.Dispose();
        }
    }
}
