//using ICSharpCode.SharpZipLib.Zip.Compression.Streams;
using Microsoft.Extensions.Configuration;

namespace CSV2iReporter; 
/// <summary>
/// 転送元のExcel列からクラスターに割り当てるデータ
/// </summary>
public class FromTo {
    /// <summary>
    /// 変換元ファイルのカラムNo（１列目＝１）
    /// </summary>
    public int ColumnNo { get; set; }

    /// <summary>
    /// 変換先帳票のシートNo
    /// </summary>
    public int SheetNo { get; set; }

    /// <summary>
    /// 変換先帳票のクラスターNo
    /// </summary>
    public int ClusterID { get; set; }

    /// <summary>
    /// 日付型の場合の入力フォーマット
    /// </summary>
    public string? FromFormat { get; set; }

    /// <summary>
    /// 日付型の場合の入力フォーマット
    /// </summary>
    public string? ToFormat { get; set; }
}

/// <summary>
/// ヘッダー部
/// </summary>
public class RemarksHeader {
    public List<KeyColumn> KeyColumn { get; set; } = null!;
    public List<RemarksDetail> Remarks { get; set; } = null!;
}
/// <summary>
/// 明細部
/// </summary>
public class RemarksDetail {
    public int Index { get; set; }
    public string? Format { get; set; }
    public string? Value { get; set; }
}

/// <summary>
/// ラベル文字置換情報
/// </summary>
public class KeyColumn {
    /// <summary>
    /// 置換文字列
    /// </summary>
    public string Key { get; set; } = null!;
    /// <summary>
    /// 置換データ取得型位置
    /// </summary>
    public int ColumnNo { get; set; }

    /// <summary>
    /// 日付文字列の入力フォーマット
    /// </summary>
    public string? FromDateFormat { get; set; }

    /// <summary>
    /// 日付文字列の出力フォーマット
    /// </summary>
    public string? ToDateFormat { get; set; }

    /// <summary>
    /// コンストラクタ
    /// </summary>
    public KeyColumn() {
    }

    /// <summary>
    /// コンストラクタ
    /// </summary>
    /// <param name="key"></param>
    /// <param name="column"></param>
    public KeyColumn(string key, int column, string? fromDateFormat, string? toDateFormat) {
        this.Key = key;
        this.ColumnNo = column;
        this.FromDateFormat = fromDateFormat;
        this.ToDateFormat = toDateFormat;
    }

    /// <summary>
    /// コンストラクタ
    /// </summary>
    /// <param name="key"></param>
    /// <param name="column"></param>
    public KeyColumn(string key, int column) : this(key,column,null,null)
    {
    }
}

///// <summary>
///// ラベル情報
///// </summary>
//public class LabelName {
//    /// <summary>
//    /// ラベル文字置換情報
//    /// </summary>
//    public List<KeyColumn> keyColumn { get; }
//    /// <summary>
//    /// ラベルベース文字列
//    /// </summary>
//    public string labelString { get; }
//    /// <summary>
//    /// コンストラクタ
//    /// </summary>
//    /// <param name="keyColumn"></param>
//    /// <param name="labelString"></param>
//    public LabelName(List<KeyColumn> keyColumn, string labelString) {
//        this.keyColumn = keyColumn;
//        this.labelString = labelString;
//    }
//}


/// <summary>
/// 設定情報
/// </summary>
public class Settings {
    /// <summary>
    /// 変換元ファイルが保存されたフォルダ名
    /// </summary>
    public string SourceFileFolder { get; }
    /// <summary>
    /// 変換元ファイル名（ワイルドカード可）
    /// </summary>
    public string SourceFileName { get; }
    /// <summary>
    /// 変換正常終了後のExcelファイルの移動先
    /// </summary>
    public string SuccessFileForder { get; }
    /// <summary>
    /// 変換異常終了後のExcelファイルの移動先
    /// </summary>
    public string ErrorFileFolder { get; }
    /// <summary>
    /// データ転送元Excelファイルの処理対象シート名
    /// </summary>
    public string ExcelSheetName { get; }
    /// <summary>
    /// データ転送元CSVファイルのエンコード
    /// </summary>
    public string CsvFileEncode { get; }
    /// <summary>
    /// Excelシートのヘッダ行数
    /// </summary>
    public int HeaderRowCount { get; }
    /// <summary>
    /// データ転送先の帳票定義ID
    /// </summary>
    public int DefTopId { get; }
    /// <summary>
    /// データ転送先の帳票定義名
    /// </summary>
    public string DefName { get; }
    /// <summary>
    /// 生成される帳票名称
    /// </summary>
    public string RepTopName { get; }
    /// <summary>
    /// ラベル情報
    /// </summary>
    public string LabelName { get; }
    /// <summary>
    /// CSV区切り文字
    /// </summary>
    public string SplitType { get; }
    /// <summary>
    /// 転送元のExcel列からクラスターに割り当てる情報
    /// </summary>
    public List<FromTo> FromTo { get; }

    /// <summary>
    /// コンストラクタ
    /// </summary>
    public Settings(IConfigurationRoot configuration) {
        var irepo = configuration.GetSection("iRepoLink");

        DefName = irepo["DefName"] ?? "";
        SourceFileFolder = irepo["SourceFileFolder"] ?? ".\\Request";
        SourceFileName = irepo["SourceFileName"] ?? "*.*";
        SuccessFileForder = irepo["SccessFileForder"] ?? ".\\Success";
        ErrorFileFolder = irepo["ErrorFileFolder"] ?? ".\\Error";
        ExcelSheetName = irepo["ExcelSheetName"] ?? "Sheet1";
        CsvFileEncode = irepo["CsvFileEncode"] ?? "UTF-8";
        HeaderRowCount = int.TryParse(irepo["HeaderRowCount"] ?? "0", out var rowCnt) ? rowCnt : 0;
        DefTopId = int.TryParse(irepo["defTopId"] ?? "", out var topId) ? topId : 0;
        RepTopName = irepo["repTopName"] ?? "{defTopName}_{datetime}";
        LabelName = irepo["LabelName"] ?? "";
        SplitType = irepo["SplitType"] ?? "AUTO";
        FromTo = irepo.GetSection("FromTo").Get<List<FromTo>>() ?? [];

    }
}
