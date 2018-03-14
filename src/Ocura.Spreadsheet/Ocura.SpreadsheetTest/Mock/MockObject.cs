using System;
using System.Collections.Generic;
using Ocura.Spreadsheet.Excel;

namespace Ocura.SpreadsheetTest.Mock
{
  public class MockObject
  {
    public MockObject(int count)
    {
      var random = new Random();

      var items = new List<Item>();
      for (var i = 0; i < count; i++)
        items.Add(new Item
        {
          Id = random.Next(),
          Name = "Name " + i,
          Description = "Description " + i,
          CheckBoolean = random.Next() % 2 > 0,
          Value = Convert.ToDecimal(random.NextDouble() * random.Next())
        });
      Items = items;

      var excelCellItems = new List<List<ExcelCell>>();
      for (var i = 0; i < count; i++)
      {
        var row = new List<ExcelCell>();
        row.Add(new ExcelCell("Name " + i, ExcelCell.Type.String));
        row.Add(new ExcelCell("Description " + i, ExcelCell.Type.String));
        row.Add(new ExcelCell(random.Next() % 2 > 0, ExcelCell.Type.Number));
        row.Add(new ExcelCell(new[] {"Google", "http://google.com"}, ExcelCell.Type.Hyperlink));
        excelCellItems.Add(row);
      }

      ExcelCellItems = excelCellItems;
    }

    public IEnumerable<IEnumerable<ExcelCell>> ExcelCellItems { get; }

    public IEnumerable<Item> Items { get; }

    public class Item
    {
      public bool CheckBoolean { get; set; }
      public string Description { get; set; }
      public int Id { get; set; }
      public string Name { get; set; }
      public decimal Value { get; set; }
    }
  }
}
