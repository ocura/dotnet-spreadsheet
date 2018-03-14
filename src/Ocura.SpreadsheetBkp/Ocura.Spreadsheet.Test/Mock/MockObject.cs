using System;
using System.Collections.Generic;

namespace Ocura.Spreadsheet.Test.Mock
{
  public class MockObject
  {
    public MockObject(int count)
    {
      var items = new List<Item>();
      var random = new Random();
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
    }

    public IEnumerable<Item> Items { get; }

    public class Item
    {
      public int Id { get; set; }
      public string Name { get; set; }
      public string Description { get; set; }
      public decimal Value { get; set; }
      public bool CheckBoolean { get; set; }
    }
  }
}
