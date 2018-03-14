namespace Ocura.Spreadsheet.Excel
{
  public class ExcelCell
  {
    /// <summary>
    ///   Enum Type
    /// </summary>
    public enum Type
    {
      String = 0,
      Number = 1,
      Hyperlink = 2 // Hyperling = new []{"Value", "URI"}
    }

    /// <summary>
    ///   Initializes a new instance of the <see cref="ExcelCell" /> class.
    /// </summary>
    public ExcelCell()
    {
    }

    /// <summary>
    ///   Initializes a new instance of the <see cref="ExcelCell" /> class.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <param name="typeId">The type identifier.</param>
    public ExcelCell(object value, Type typeId)
    {
      Value = value;
      TypeId = typeId;
    }

    /// <summary>
    ///   Gets or sets the type identifier.
    /// </summary>
    /// <value>
    ///   The type identifier.
    /// </value>
    public Type TypeId { get; set; }

    /// <summary>
    ///   Gets or sets the value.
    /// </summary>
    /// <value>
    ///   The value.
    /// </value>
    public object Value { get; set; }
  }
}
