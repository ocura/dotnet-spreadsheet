using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Newtonsoft.Json;
using Ocura.Spreadsheet.Csv;
using Ocura.Spreadsheet.Excel;
using Ocura.SpreadsheetTest.Mock;
using Xunit;

namespace Ocura.SpreadsheetTest
{
  public class SpreadsheetTest
  {
    private readonly MockObject _mock = new MockObject(25);
    private readonly string _applicationPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

    [Fact]
    public void GenerateCsvFromObjectList()
    {
      Spreadsheet.Spreadsheet spreadsheet = new Csv();
      spreadsheet.SetData(_mock.Items);
      var file = spreadsheet.GenerateFile();
      File.WriteAllBytes(_applicationPath + @"\GenerateCsvFromObjectList.csv", file);
    }

    [Fact]
    public void GenerateExcelFromExcelCellListWithStylesheet()
    {
      var spreadsheet = new Excel();
      spreadsheet.Stylesheet = ExcelStylesheet.Default;
      spreadsheet.SetStylesheet(0, 1);
      spreadsheet.SetStylesheet(7, 1);

      var header = new List<string> {"Meu id", "Meu nome", "minha desc", "meu valor"};
      spreadsheet.SetDataByExcelCell(header, _mock.ExcelCellItems);
      var file = spreadsheet.GenerateFile();
      File.WriteAllBytes(_applicationPath + @"\GenerateExcelFromExcelCellListWithStylesheet.xlsx", file);
    }

    [Fact]
    public void GenerateExcelFromObjectList()
    {
      Spreadsheet.Spreadsheet spreadsheet = new Excel();
      var header = new List<string> {"Meu id", "Meu nome", "minha desc", "meu valor", "Oi?"};
      spreadsheet.SetData(header, _mock.Items);
      var file = spreadsheet.GenerateFile();
      File.WriteAllBytes(_applicationPath + @"\GenerateExcelFromObjectList.xlsx", file);
    }

    [Fact]
    public void GenerateExcelFromObjectListWithStylesheet()
    {
      var spreadsheet = new Excel();
      spreadsheet.Stylesheet = ExcelStylesheet.Default;
      spreadsheet.SetStylesheet(0, 1);
      spreadsheet.SetStylesheet(7, 1);

      var header = new List<string> {"Meu id", "Meu nome", "minha desc", "meu valor", "Oi?"};
      spreadsheet.SetData(header, _mock.Items);
      var file = spreadsheet.GenerateFile();
      File.WriteAllBytes(_applicationPath + @"\GenerateExcelFromObjectListWithStylesheet.xlsx", file);
    }

    [Fact]
    public void SpreadsheetReadFile()
    {
      var excelMapper = new ExcelMapper();
      excelMapper.Add("A", "HotelKey", ExcelMapper.Type.String);
      excelMapper.Add("B", "HotelName", ExcelMapper.Type.String);
      excelMapper.Add("C", null, ExcelMapper.Type.String);
      excelMapper.Add("D", "LocationType", ExcelMapper.Type.String);
      excelMapper.Add("E", "Address_1", ExcelMapper.Type.String);
      excelMapper.Add("F", "Address_2", ExcelMapper.Type.String);
      excelMapper.Add("G", "Address_3", ExcelMapper.Type.String);
      excelMapper.Add("H", "Address_4", ExcelMapper.Type.String);
      excelMapper.Add("I", "Phone", ExcelMapper.Type.String);
      excelMapper.Add("J", "Fax", ExcelMapper.Type.String);
      excelMapper.Add("K", "Email", ExcelMapper.Type.String);
      excelMapper.Add("L", "Website", ExcelMapper.Type.String);
      excelMapper.Add("M", "StarRating", ExcelMapper.Type.Int);
      excelMapper.Add("N", "Category", ExcelMapper.Type.String);
      excelMapper.Add("O", "Latitude", ExcelMapper.Type.Double);
      excelMapper.Add("P", "Longitude", ExcelMapper.Type.Double);
      excelMapper.Add("Q", "CityCode", ExcelMapper.Type.String);
      excelMapper.Add("R", "CityName", ExcelMapper.Type.String);
      excelMapper.Add("S", "CountryCode", ExcelMapper.Type.String);
      excelMapper.Add("T", null, ExcelMapper.Type.String);
      excelMapper.Add("U", "CountryName", ExcelMapper.Type.String);

      JsonConvert.SerializeObject(excelMapper);

      var fileName = @"E:\Source\Netbiis-Git\dotnet-spreadsheet\src\Netbiis.Spreadsheet\Netbiis.SpreadsheetTest\Files\HotelProvider.xlsx";
      var fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
      var spreadsheet = new Excel();
      var hotelProvider = spreadsheet.ReadFile<HotelProvider>(fileStream, excelMapper, true);
    }
  }
}
