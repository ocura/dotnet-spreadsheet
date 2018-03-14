using System.IO;
using System.Reflection;
using Ocura.Spreadsheet.Test.Mock;
using Xunit;

namespace Ocura.Spreadsheet.Test
{
  public class SpreadsheetTest
  {
    private readonly MockObject _mock = new MockObject(25);
    private readonly string _applicationPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

    [Fact]
    public void GenerateCsvFromObjectList()
    {
      Spreadsheet spreadsheet = new Csv();
      spreadsheet.SetData(_mock.Items);
      var file = spreadsheet.GenerateFile();
      File.WriteAllBytes(_applicationPath + @"\file.csv", file);
    }

    [Fact]
    public void GenerateExcelFromObjectList()
    {
      Spreadsheet spreadsheet = new Excel();
      spreadsheet.SetData(_mock.Items);
      var file = spreadsheet.GenerateFile();
      File.WriteAllBytes(_applicationPath + @"\file.xlsx", file);
    }
  }
}
