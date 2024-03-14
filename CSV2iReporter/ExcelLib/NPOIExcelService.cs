using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelLib;
/// <summary>
/// コンストラクタ
/// </summary>
/// <param name="workbook"></param>
public class NPOIWorkBook(XSSFWorkbook workbook) : IWorkBook {
    private readonly XSSFWorkbook workbook = workbook;

    public int NumberOfSheets => workbook?.NumberOfSheets ?? 0;

    public ISheet GetWorkSheet(string name) =>
        new NPOIWorkSheet(workbook.GetSheet(name));

    public ISheet GetSheetAt(int index) =>
        new NPOIWorkSheet(workbook.GetSheetAt(index - 1));

    public ISheet CreateSheet() {
        var s = workbook.CreateSheet();
        workbook.Add(s);
        return new NPOIWorkSheet(s);
    }

    public void Close() =>
        workbook.Close();

    public void Dispose() {
        workbook?.Dispose();
        GC.SuppressFinalize(this);
    }

    public void Write(string fnam) {
        using var fs = new FileStream(fnam, FileMode.Create);
        workbook?.Write(fs, false);
    }
}

/// <summary>
/// コンストラクタ
/// </summary>
/// <param name="sheet"></param>
public class NPOIWorkSheet(NPOI.SS.UserModel.ISheet sheet) : ISheet {
    private readonly NPOI.SS.UserModel.ISheet sheet = sheet;

    public bool IsSheetAvailable => sheet != null;
    public string SheetName => sheet.SheetName;

    public IRow? GetRow(int row, bool force = false) {
        var r = sheet.GetRow(row - 1);       
        if (r == null) {
            if (force) {
                r = sheet.CreateRow(row - 1);
            } else {
                return null;
            }
        }
        return new NPOIRow(r);
    }

    public string GetStringValue(int row, int column) {
        var result = "";
        var rowCell = GetRow(row);
        if (rowCell != null) {
            var clmCell = rowCell.GetCell(column);
            if (clmCell != null) {
                result = clmCell.GetStringValue();
            }
        }

        return result;
    }

    public string GetStringValue(RowCol rc) => GetStringValue(rc.Row, rc.Col);

    public void SetCellValue(int row, int clm, object value) {
        var rowCell = GetRow(row);
        if (rowCell is not null) {
            var clmCell = rowCell.GetCell(clm);
            clmCell?.SetValue($"{value}");
        }
    }

    public void CreateFreezePane(int col, int row) {
        sheet.CreateFreezePane(col, row);
    }

    public int GetLastRowNum() => sheet.LastRowNum + 1;
}

public class NPOIRow(NPOI.SS.UserModel.IRow row) : IRow {
    readonly NPOI.SS.UserModel.IRow row = row;

    public ICell? GetCell(int column, bool force = false) {
        var cell = row.GetCell(column - 1);
        if (cell == null) {
            if (force) {
                cell = row.CreateCell(column - 1);
            } else {
                return null;
            }
        }
        return new NPOICell(cell);
    }

    public int GetLastCellNum() => this.row.LastCellNum;
}

public class NPOICell(NPOI.SS.UserModel.ICell cell) : ICell {
    readonly NPOI.SS.UserModel.ICell cell = cell;

    public string StringCellValue => cell.StringCellValue;

    public double NumericCellValue => cell.NumericCellValue;

    public bool BooleanCellValue => cell.BooleanCellValue;

    public CellType CellType => cell.CellType switch {
        NPOI.SS.UserModel.CellType.String => CellType.String,
        NPOI.SS.UserModel.CellType.Numeric => CellType.Numeric,
        NPOI.SS.UserModel.CellType.Boolean => CellType.Boolean,
        NPOI.SS.UserModel.CellType.Formula => cell.CachedFormulaResultType switch {
            NPOI.SS.UserModel.CellType.String => CellType.String,
            NPOI.SS.UserModel.CellType.Numeric => CellType.Numeric,
            NPOI.SS.UserModel.CellType.Boolean => CellType.Boolean,
            _ => CellType.String
        },
        _ => CellType.String
    };

    public string GetStringValue() =>
        CellType switch {
            CellType.String => StringCellValue,
            CellType.Numeric => $"{GetNumOrDate(cell)}",
            CellType.Boolean => $"{BooleanCellValue}",
            CellType.Formula => StringCellValue,
            _ => StringCellValue
        };

    public static string GetNumOrDate(NPOI.SS.UserModel.ICell cell) {
        string? result;
        if (DateUtil.IsCellDateFormatted(cell)) {
            // 日付型
            // 本来はスタイルに合わせてフォーマットすべきだが、
            // うまく表示できないケースが若干見られたので固定のフォーマットとして取得
            result = cell.DateCellValue.ToString("yyyy/MM/dd HH:mm:ss");
        } else {
            // 数値型
            result = cell.NumericCellValue.ToString();
        }
        return result;
    }

    public void SetValue(string value) => cell.SetCellValue(value);
}

public class NPOIExcelService : IExcelService {

    public IWorkBook Open() {
        var book = new XSSFWorkbook();
        return new NPOIWorkBook(book);
    }
    public IWorkBook Open(string path) {
        var book = new XSSFWorkbook(path);
        return new NPOIWorkBook(book);
    }
}
