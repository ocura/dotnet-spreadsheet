using DocumentFormat.OpenXml.Spreadsheet;

namespace Ocura.Spreadsheet.Excel
{
  public static class ExcelStylesheet
  {
    public static Stylesheet Default => SetDefault();

    private static Stylesheet SetDefault()
    {
      //Fonts
      var fonts = new Fonts();

      var font0 = new Font();
      var fontSize0 = new FontSize {Val = 10};
      font0.AppendChild(fontSize0);

      var font1 = new Font();
      var fontSize1 = new FontSize {Val = 12};
      var fontBold1 = new Bold();
      var fontColor1 = new Color {Rgb = "FFFFFF"};
      font1.AppendChild(fontSize1);
      font1.AppendChild(fontBold1);
      font1.AppendChild(fontColor1);

      fonts.AppendChild(font0);
      fonts.AppendChild(font1);

      // Alignments
      var alignmentCenter = new Alignment();
      alignmentCenter.Horizontal = HorizontalAlignmentValues.Center;

      //Fill
      var fills = new Fills();

      var fill0 = new Fill();
      var patternFill0 = new PatternFill {PatternType = PatternValues.None};
      fill0.Append(patternFill0);

      var fillSkip = new Fill(); // Not valid needs to skip
      var patternFillSkip = new PatternFill {PatternType = PatternValues.None};
      fillSkip.Append(patternFillSkip);

      var fill1 = new Fill();
      var patternFill1 = new PatternFill {PatternType = PatternValues.Solid};
      var foregroundColor1 = new ForegroundColor {Rgb = "00000000"};
      var backgroundColor1 = new BackgroundColor {Indexed = 64U};
      patternFill1.Append(foregroundColor1);
      patternFill1.Append(backgroundColor1);
      fill1.Append(patternFill1);

      fills.Append(fill0);
      fills.Append(fillSkip);
      fills.Append(fill1);

      //Border
      var borders = new Borders();

      var border0 = new Border();
      var border1 = new Border();

      borders.AppendChild(border0);
      borders.AppendChild(border1);

      var cellFormats = new CellFormats(
        new CellFormat {FontId = 0, FillId = 0, BorderId = 0, ApplyFill = true}, // Index 0 
        new CellFormat {FontId = 1, FillId = 2, BorderId = 1, ApplyFill = true, Alignment = alignmentCenter} // Index 1
      );

      return new Stylesheet(fonts, fills, borders, cellFormats);
    }
  }
}
