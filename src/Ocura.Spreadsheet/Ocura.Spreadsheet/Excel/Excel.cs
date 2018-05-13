using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using AutoMapper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ocura.Helper;

namespace Ocura.Spreadsheet.Excel
{
  public class Excel : Spreadsheet
  {
    /// <summary>
    ///   The row stylesheet
    /// </summary>
    private readonly List<KeyValuePair<int, uint?>> _rowStylesheet;

    /// <summary>
    ///   Initializes a new instance of the <see cref="Excel" /> class.
    /// </summary>
    public Excel()
    {
      _rowStylesheet = new List<KeyValuePair<int, uint?>>();
    }

    /// <summary>
    ///   Gets or sets the stylesheet.
    /// </summary>
    /// <value>
    ///   The stylesheet.
    /// </value>
    public Stylesheet Stylesheet { get; set; }

    /// <summary>
    ///   Converts the excel cell to cell.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <returns></returns>
    private static Cell ConvertExcelCellToCell(ExcelCell value)
    {
      var valueType = value.TypeId;

      switch (valueType)
      {
        case ExcelCell.Type.Hyperlink:
          var convValue = (string[]) value.Value;
          var cell = new Cell();
          cell.DataType = CellValues.String;
          var cellformula = new CellFormula();
          cellformula.Text = "HYPERLINK(\"" + convValue[1] + "\", \"" + convValue[0] + "\")";
          var cellValue = new CellValue(convValue[0]);
          cell.AppendChild(cellformula);
          cell.AppendChild(cellValue);
          return cell;
        default:
          return ConvertObjectToCell(value.Value);
      }
    }

    /// <summary>
    ///   Converts the object to cell.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <returns></returns>
    private static Cell ConvertObjectToCell(object value)
    {
      var cell = new Cell();
      if (value == null)
      {
        cell.DataType = CellValues.String;
        var cellValue = new CellValue("");
        cell.AppendChild(cellValue);
        return cell;
      }

      var objType = value.GetType();

      if (objType == typeof(ExcelCell))
        return ConvertExcelCellToCell((ExcelCell) value);

      if (objType == typeof(decimal) || objType == typeof(int))
      {
        cell.DataType = CellValues.Number;
        cell.CellValue = new CellValue(value.ToString());
      }
      else
      {
        cell.DataType = CellValues.String;
        var cellValue = new CellValue(value.ToString() ?? "");
        cell.AppendChild(cellValue);
      }

      return cell;
    }

    /// <summary>
    ///   Generates the file.
    /// </summary>
    /// <returns></returns>
    public override byte[] GenerateFile()
    {
      var values = new List<IEnumerable<object>>();
      values.Add(Header);
      values.AddRange(Body);

      var memory = new MemoryStream();
      using (var spreadsheetDocument = SpreadsheetDocument.Create(memory, SpreadsheetDocumentType.Workbook))
      {
        var workbookPart = spreadsheetDocument.AddWorkbookPart();

        if (Stylesheet != null)
        {
          // Adding style
          var stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
          stylePart.Stylesheet = Stylesheet;
          stylePart.Stylesheet.Save();
        }

        var sheetName = "Document";
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();

        worksheetPart.Worksheet = new Worksheet(sheetData);
        workbookPart.Workbook = new Workbook();
        workbookPart.Workbook.AppendChild(new Sheets());

        var sheet = new Sheet
        {
          Id = workbookPart.GetIdOfPart(worksheetPart),
          SheetId = 1,
          Name = sheetName
        };

        for (var i = 0; i < values.Count; i++)
        {
          var value = values[i];
          var row = new Row {RowIndex = (uint) i + 1};
          foreach (var col in value)
          {
            var cell = ConvertObjectToCell(col);
            var styleIndex = _rowStylesheet.FirstOrDefault(a => a.Key == i).Value;
            if (styleIndex != null)
              cell.StyleIndex = styleIndex;
            row.AppendChild(cell);
          }

          sheetData.AppendChild(row);
        }

        workbookPart.Workbook.Sheets.AppendChild(sheet);
        workbookPart.Workbook.Save();
      }

      TextWriter writer = new StreamWriter(memory);
      writer.Flush();
      memory.Position = 0;
      var file = memory.ToArray();
      memory.Dispose();

      return file;
    }

    /// <summary>
    ///   Reads the file.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="excelFile">The excel file.</param>
    /// <param name="excelMapper">The excel mapper.</param>
    /// <param name="hasTitle">if set to <c>true</c> [has title].</param>
    /// <returns></returns>
    public IEnumerable<T> ReadFile<T>(FileStream excelFile, ExcelMapper excelMapper, bool hasTitle) where T : class
    {
      using (var doc = SpreadsheetDocument.Open(excelFile, false))
      {
        var config = new MapperConfiguration(cfg => cfg.CreateMap<List<dynamic>, IEnumerable<T>>());
        var mapper = config.CreateMapper();

        var workbookPart = doc.WorkbookPart;
        var sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
        var sst = sstpart.SharedStringTable;

        var worksheetPart = workbookPart.WorksheetParts.First();
        var sheet = worksheetPart.Worksheet;

        var rows = sheet.Descendants<Row>();

        var retObject = new List<ExpandoObject>();
        var rowsArray = rows.ToArray();

        for (var r = hasTitle ? 1 : 0; r < rowsArray.Length; r++)
        {
          var row = rowsArray[r];
          var cs = row.Elements<Cell>().ToArray();
          var dataRow = new ExpandoObject() as IDictionary<string, object>;
          for (var j = 0; j < cs.Length; j++)
          {
            var c = cs[j];
            object stgValue;

            var cellReferenceLetter = new Regex("[A-Za-z]+").Match(c.CellReference).Value;
            var excelMap = excelMapper.Items.FirstOrDefault(a => a.From == cellReferenceLetter);

            if (excelMap == null) continue;

            if (c.DataType != null && c.DataType == CellValues.SharedString)
            {
              var ssid = int.Parse(c.CellValue.Text);
              var str = sst.ChildElements[ssid].InnerText;
              stgValue = str;
            }
            else
            {
              stgValue = c.CellValue?.Text ;
            }

            object convertedValue;

            switch (excelMap.Type)
            {
              case ExcelMapper.Type.Double:
                convertedValue = Convert.ToDouble(stgValue ?? 0);
                break;
              case ExcelMapper.Type.Int:
                convertedValue = Convert.ToInt32(stgValue ?? 0);
                break;
              case ExcelMapper.Type.String:
                convertedValue = stgValue;
                break;
              default:
                convertedValue = stgValue;
                break;
            }

            var dataField = (IDictionary<string, object>) DynamicObjectHelper
              .GenerateNestedObject(excelMap.To ?? excelMap.From, convertedValue);

            dataRow = DynamicObjectHelper
              .AddIDictionary((ExpandoObject) dataRow, (ExpandoObject) dataField);
          }
          retObject.Add((ExpandoObject) dataRow);
        }

        return retObject.Select(c => mapper.Map<T>(c));
      }
    }

    /// <summary>
    ///   Sets the data by ExcelCell.
    /// </summary>
    /// <param name="header">The header.</param>
    /// <param name="body">The body.</param>
    /// <exception cref="ArgumentNullException"></exception>
    public void SetDataByExcelCell(IEnumerable<string> header, IEnumerable<IEnumerable<ExcelCell>> body)
    {
      Header = header;
      Body = body;
    }

    /// <summary>
    ///   Sets the stylesheet.
    /// </summary>
    /// <param name="row">The row.</param>
    /// <param name="styleIndex">Index of the style.</param>
    public void SetStylesheet(int row, uint styleIndex)
    {
      _rowStylesheet.Add(new KeyValuePair<int, uint?>(row, styleIndex));
    }
  }
}
