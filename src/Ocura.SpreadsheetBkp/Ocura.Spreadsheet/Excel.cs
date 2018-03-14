using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ocura.Spreadsheet
{
  public class Excel : Spreadsheet
  {
    /// <summary>
    /// Generates the file.
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
            row.AppendChild(ConvertObjectToCell(col));
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
    /// Converts the object to cell.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <returns></returns>
    private static Cell ConvertObjectToCell(object value)
    {
      var objType = value.GetType();
      var cell = new Cell();
      if (objType == typeof(decimal))
      {
        cell.DataType = CellValues.Number;
        cell.CellValue = new CellValue(value.ToString());
      }
      else if (objType == typeof(int))
      {
        cell.DataType = CellValues.Number;
        cell.CellValue = new CellValue(value.ToString());
      }
      else
      {
        cell.DataType = CellValues.InlineString;
        var inlineString = new InlineString();
        var text = new Text {Text = value.ToString()};
        inlineString.AppendChild(text);
        cell.AppendChild(inlineString);
      }

      return cell;
    }
  }
}
