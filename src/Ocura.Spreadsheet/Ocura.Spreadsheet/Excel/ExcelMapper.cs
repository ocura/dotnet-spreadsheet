using System;
using System.Collections.Generic;

namespace Ocura.Spreadsheet.Excel
{
  public class ExcelMapper
  {
    public enum Type
    {
      String,
      Int,
      Double
    }

    public ExcelMapper()
    {
      Items = new List<Item>();
    }

    public List<Item> Items { get; set; }

    public void Add(string from, string to, string type)
    {
      var item = new Item
      {
        From = from,
        To = to
      };

      if (string.Equals(type, Type.String.ToString(), StringComparison.CurrentCultureIgnoreCase))
        item.Type = Type.String;
      else if (string.Equals(type, Type.Int.ToString(), StringComparison.CurrentCultureIgnoreCase))
        item.Type = Type.Int;
      else if (string.Equals(type, Type.Double.ToString(), StringComparison.CurrentCultureIgnoreCase))
        item.Type = Type.Double;
      else
        throw new ArgumentOutOfRangeException(nameof(type), type, null);

      Items.Add(item);
    }

    public void Add(string from, string to, Type type)
    {
      var item = new Item
      {
        From = from,
        To = to,
        Type = type
      };
      Items.Add(item);
    }

    public class Item
    {
      public string From { get; set; }
      public string To { get; set; }
      public Type Type { get; set; }
    }
  }
}
