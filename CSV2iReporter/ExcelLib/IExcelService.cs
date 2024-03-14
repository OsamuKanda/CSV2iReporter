using NPOI.SS.UserModel;

namespace ExcelLib; 
public interface IExcelService {
    public IWorkBook Open(string path);
    public IWorkBook Open();
}

public interface IWorkBook : IDisposable {
    public int NumberOfSheets { get; }
    public ISheet GetWorkSheet(string name);
    /// <summary>
    /// 
    /// </summary>
    /// <param name="index">1で始まる</param>
    /// <returns></returns>
    public ISheet GetSheetAt(int index);

    public ISheet CreateSheet();

    public void Close();

    public void Write(string fnam);
}

public interface ISheet {
    /// <summary>
    /// 
    /// </summary>
    /// <param name="row">1で始まる</param>
    /// <param name="column">1で始まる</param>
    /// <returns></returns>
    public string GetStringValue(int row, int column);

    public string GetStringValue(RowCol rc);
    public bool IsSheetAvailable { get; }
    public string SheetName { get; }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="row">1で始まる</param>
    /// <param name="force">なければ生成して返す</param>
    /// <returns></returns>
    public IRow? GetRow(int row, bool force = false);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="row">1で始まる</param>
    /// <param name="clm">1で始まる</param>
    /// <param name="value"></param>
    public void SetCellValue(int row, int clm, object value);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="col"></param>
    /// <param name="row"></param>
    public void CreateFreezePane(int col, int row);

    public int GetLastRowNum();
}
public interface IRow {
    /// <summary>
    /// 
    /// </summary>
    /// <param name="column">1で始まる</param>
    /// <param name="force">なければ生成して返す</param>
    /// <returns></returns>
    public ICell? GetCell(int column, bool force = false);

    /// <summary>
    /// 
    /// </summary>
    /// <returns>1で始まる</returns>
    public int GetLastCellNum();
 }
public interface ICell {
    string StringCellValue { get; }
    double NumericCellValue { get; }
    bool BooleanCellValue { get; }

    CellType CellType { get; }

    /// <summary>
    /// string値に変換して返す
    /// </summary>
    /// <returns></returns>
    string GetStringValue();

    void SetValue(string value);
}

public enum CellType {
    String,
    Numeric,
    Boolean,
    Formula
}

public class RowCol {
    public int Row { get; set; } = -1;
    public int Col { get; set; } = -1;
}
