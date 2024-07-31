//using ICSharpCode.SharpZipLib.Zip.Compression.Streams;
using Microsoft.Extensions.Configuration;
using Serilog;
using System.Text.Json.Serialization;
using System.Text;
using System.Runtime;


namespace CSV2iReporter; 
/// <summary>
/// 転送元の列からクラスターに割り当てるデータ
/// </summary>
public class FromTo {

    /// <summary>
    /// 引数無しのコンストラクタ
    /// </summary>
    public FromTo() { }


    /// <summary>
    /// 引数をすべて指定するコンストラクタ
    /// </summary>
    /// <param name="columnNo"></param>
    /// <param name="sheetNo"></param>
    /// <param name="clusterID"></param>
    /// <param name="fromDateFormat"></param>
    /// <param name="toDateFormat"></param>
    public FromTo(int columnNo, int? sheetNo, int? clusterID, int? remarksNo, int? systemKeyNo, string? fromDateFormat, string? toDateFormat) {
        ColumnNo = columnNo;
        SheetNo = sheetNo;
        ClusterID = clusterID;
        RemarksNo = remarksNo;
        SystemKeyNo = systemKeyNo;
        FromDateFormat = fromDateFormat;
        ToDateFormat = toDateFormat;
    }
    //FromDateFormatとToDateFormatを指定しないコンストラクタ
    public FromTo(int columnNo, int? sheetNo, int? clusterID) : this(columnNo, sheetNo, clusterID, null, null, null, null) { }

    //ToDateFormatを指定しないコンストラクタ
    public FromTo(int columnNo, int? sheetNo, int? clusterID, string fromDateFormat) : this(columnNo, sheetNo, clusterID, null, null, fromDateFormat, null) { }

    /// <summary>
    /// 変換元ファイルのカラム名または番号（列名または1～の名称）
    /// </summary>
    public int ColumnNo { get; set; }

    /// <summary>
    /// 変換先帳票のシートNo
    /// </summary>
    public int? SheetNo { get; set; }

    /// <summary>
    /// 変換先帳票のクラスターNo
    /// </summary>
    public int? ClusterID { get; set; }

    /// <summary>
    /// 帳票備考番号
    /// </summary>
    public int? RemarksNo { get; set; }

    /// <summary>
    /// システムキー番号
    /// </summary>
    public int? SystemKeyNo { get; set; }

    /// <summary>
    /// 日付型の場合の入力フォーマット
    /// </summary>
    public string? FromDateFormat { get; set; }

    /// <summary>
    /// 日付型の場合の入力フォーマット
    /// </summary>
    public string? ToDateFormat { get; set; }

}

/// <summary>
/// 設定情報
/// </summary>
public class Settings {
    /// <summary>
    /// 変換元ファイルが保存されたフォルダ名
    /// </summary>
    public string SourceFolder { get; }
    /// <summary>
    /// 変換元ファイル名（ワイルドカード可）
    /// </summary>
    public string SourceFileName { get; }
    /// <summary>
    /// 変換正常終了後のファイルの移動先フォルダ名
    /// </summary>
    public string SuccessFileMoveForder { get; }
    /// <summary>
    /// 変換異常終了後のファイルの移動先フォルダ名
    /// </summary>
    public string ErrorFileMoveFolder { get; }
    /// <summary>
    /// データ転送元CSVファイルのエンコード
    /// </summary>
    public Encoding Encode { get; }
    /// <summary>
    /// ヘッダ行数
    /// </summary>
    public int HeaderRowCount { get; }
    /// <summary>
    /// データ転送先の帳票定義ID
    /// </summary>
    public int? DefTopId { get; }
    /// <summary>
    /// データ転送先の帳票定義名
    /// </summary>
    public string? DefTopName { get; }
    /// <summary>
    /// 生成される帳票名称
    /// </summary>
    public string RepTopName { get; }
    /// <summary>
    /// ラベル情報
    /// </summary>
    public string? LabelName { get; }
    /// <summary>
    /// CSV区切り文字
    /// </summary>
    public string SeparateChar { get; }
    /// <summary>
    /// 転送元の列からクラスターに割り当てる情報
    /// </summary>
    public List<FromTo> FromTo { get; }= new List<FromTo>();

    /// <summary>
    /// コンストラクタ
    /// </summary>
    public Settings(IConfigurationRoot configuration) {
        var irepo = configuration.GetSection("Convert");

        // 帳票定義ID
        DefTopId = int.TryParse(irepo[$"{nameof(DefTopId)}"], out var defTopId) ? defTopId : null;
        // 帳票定義名
        DefTopName = irepo[$"{nameof(DefTopName)}"];
        // 読込ファイル保存フォルダ
        SourceFolder = irepo[$"{nameof(SourceFolder)}"] ?? @".\Request";
        // 保存ファイル名
        SourceFileName = irepo[$"{nameof(SourceFileName)}"] ?? "*.csv";
        // 処理成功ファイル保存フォルダ
        SuccessFileMoveForder = irepo[$"{nameof(SuccessFileMoveForder)}"] ?? @".\Success";
        // 処理エラーファイル保存フォルダ
        ErrorFileMoveFolder = irepo[$"{nameof(ErrorFileMoveFolder)}"] ?? @".\Error";
        // 文字エンコード
        switch (irepo[$"{nameof(Encode)}"]??"utf-8".ToLower()) {
            case "sjis" or "shift-jis" or "shift_jis":
                Encode = Encoding.GetEncoding("shift-jis");
                break;
            case "utf8" or "utf-8" or "utf_8":
                Encode = Encoding.UTF8;
                break;
            case "unicode" or "utf-16" or "utf_16":
                Encode = Encoding.Unicode;
                break;
            case "euc-jp" or "euc_jp":
                Encode = Encoding.GetEncoding("euc-jp");
                break;
            case "iso-2022-jp" or "iso_2022_jp" or"jis":
                Encode = Encoding.GetEncoding("iso-2022-jp");
                break;
            default:
                Encode = Encoding.UTF8;
                break;
        }
        // CSVファイル内のヘッダ行数
        HeaderRowCount = int.TryParse(irepo[$"{nameof(HeaderRowCount)}"], out var rowCnt) ? rowCnt : 0;
        // 生成されるレポート名
        RepTopName = irepo[$"{nameof(RepTopName)}"] ?? "{defTopName}_{datetime}";
        // 生成されるラベル名
        LabelName = irepo[$"{nameof(LabelName)}"];
        // CSV区切り文字
        SeparateChar = irepo[$"{nameof(SeparateChar)}"] ?? "AUTO";

        // 変換セクション
        foreach (var x in irepo.GetSection($"{nameof(FromTo)}").GetChildren() ) {
            var d = new FromTo();

            // 列番号（必須）
            d.ColumnNo = int.TryParse(x["ColumnNo"], out var columnNo) ? columnNo : -1;
            // シート番号
            d.SheetNo = int.TryParse(x["SheetNo"], out var sheetNo) ? sheetNo : null;
            // クラスターID
            d.ClusterID = int.TryParse(x["ClusterID"], out var clusterID) ? clusterID : null;
            // 備考番号（１～１０）
            d.RemarksNo = int.TryParse(x["RemarksNo"], out var remarksNo) ? remarksNo : null;
            // システムキー（１～４）
            d.SystemKeyNo = int.TryParse(x["SystemKeyNo"], out var systemKeyNo) ? systemKeyNo : null;
            // 変換元日付フォーマット
            d.FromDateFormat = x["FromDateFormat"];
            // 変換先日付フォーマット
            if( d.FromDateFormat is not null ) {
                d.ToDateFormat = x["ToDateFormat"];
            }

            // リストに追加
            FromTo.Add(d);
        }
    }
}
