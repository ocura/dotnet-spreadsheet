using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Ocura.Spreadsheet
{
  public class Csv : Spreadsheet
  {
    /// <summary>
    /// Generates the file.
    /// </summary>
    /// <returns></returns>
    public override byte[] GenerateFile()
    {
      const char separator = ';';

      var values = new List<IEnumerable<string>>();
      values.Add(Header.Select(a => a.ToString()));
      values.AddRange(Body.Select(a => a.Select(b => b.ToString())));

      var fileContent = "";
      foreach (var row in values)
      {
        fileContent = row.Aggregate(fileContent,
          (current, column) => current + ("\"" + column?.Replace("\"", "\\\"") + "\"" + separator));

        fileContent = fileContent.Trim(separator);
        fileContent += Environment.NewLine;
      }

      byte[] file;
      using (var memory = new MemoryStream())
      {
        TextWriter writer = new StreamWriter(memory);
        writer.Write(fileContent);
        writer.Flush();
        memory.Position = 0;
        file = memory.ToArray();
      }

      return file;
    }
  }
}
