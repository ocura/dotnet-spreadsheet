using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Ocura.Spreadsheet.Csv
{
  public class Csv : Spreadsheet
  {
    /// <summary>
    ///   Initializes a new instance of the <see cref="Csv" /> class.
    /// </summary>
    public Csv()
    {
      Separator = ',';
    }

    /// <summary>
    ///   Gets or sets the separator.
    /// </summary>
    /// <value>
    ///   The separator.
    /// </value>
    public char Separator { get; set; }

    /// <summary>
    ///   Generates the file.
    /// </summary>
    /// <returns></returns>
    public override byte[] GenerateFile()
    {
      var values = new List<IEnumerable<string>>();
      values.Add(Header.Select(a => a.ToString()));
      values.AddRange(Body.Select(a => a.Select(b => b.ToString())));

      var fileContent = "";
      foreach (var row in values)
      {
        fileContent = row.Aggregate(fileContent,
          (current, column) => current + ("\"" + column?.Replace("\"", "\\\"") + "\"" + Separator));

        fileContent = fileContent.Trim(Separator);
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
