using System;
using System.Collections.Generic;
using System.Linq;

namespace Ocura.Spreadsheet
{
  public abstract class Spreadsheet
  {
    /// <summary>
    /// Gets the header.
    /// </summary>
    /// <value>
    /// The header.
    /// </value>
    public IEnumerable<object> Header { get; private set; }

    /// <summary>
    /// Gets the body.
    /// </summary>
    /// <value>
    /// The body.
    /// </value>
    public IEnumerable<IEnumerable<object>> Body { get; private set; }

    /// <summary>
    /// Sets the data.
    /// </summary>
    /// <param name="header">The header.</param>
    /// <param name="body">The body.</param>
    /// <exception cref="ArgumentNullException"></exception>
    public void SetData(IEnumerable<string> header, IEnumerable<object> body)
    {
      if (body == null) throw new ArgumentNullException();

      var type = body.First().GetType();
      var props = type.GetProperties();

      if (header == null) header = props.Select(prop => prop.Name).ToList();

      Header = header;
      Body = body.Select(obj => props.Select(prop => prop.GetValue(obj, null))).ToList();
    }

    /// <summary>
    /// Sets the data.
    /// </summary>
    /// <param name="header">The header.</param>
    /// <param name="body">The body.</param>
    /// <exception cref="ArgumentNullException">
    /// </exception>
    public void SetData(IEnumerable<string> header, IEnumerable<IEnumerable<string>> body)
    {
      Header = header ?? throw new ArgumentNullException();
      Body = body ?? throw new ArgumentNullException();
    }

    /// <summary>
    /// Sets the data.
    /// </summary>
    /// <param name="data">The data.</param>
    /// <exception cref="ArgumentNullException">
    /// </exception>
    public void SetData(IEnumerable<IEnumerable<string>> data)
    {
      Header = data?.First() ?? throw new ArgumentNullException();
      Body = data.Skip(1) ?? throw new ArgumentNullException();
    }

    /// <summary>
    /// Sets the data.
    /// </summary>
    /// <param name="data">The data.</param>
    public void SetData(IEnumerable<object> data)
    {
      SetData(null, data);
    }

    /// <summary>
    /// Generates the file.
    /// </summary>
    /// <returns></returns>
    public abstract byte[] GenerateFile();
  }
}
