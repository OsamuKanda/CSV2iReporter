//using ICSharpCode.SharpZipLib.Zip.Compression.Streams;
using Microsoft.Extensions.Configuration;
using System.Text.Json.Serialization;

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
    public FromTo(int columnNo, int sheetNo, int clusterID, string? fromDateFormat, string? toDateFormat) {
        ColumnNo = columnNo;
        SheetNo = sheetNo;
        ClusterID = clusterID;
        FromDateFormat = fromDateFormat;
        ToDateFormat = toDateFormat;
    }
    //FromDateFormatとToDateFormatを指定しないコンストラクタ
    public FromTo(int columnNo, int sheetNo, int clusterID) : this(columnNo, sheetNo, clusterID, null, null) { }

    //ToDateFormatを指定しないコンストラクタ
    public FromTo(int columnNo, int sheetNo, int clusterID, string fromDateFormat) : this(columnNo, sheetNo, clusterID, fromDateFormat, null) { }

    /// <summary>
    /// 変換元ファイルのカラム名または番号（列名または1～の名称）
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
    public string Encode { get; }
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
    public string RepName { get; }
    /// <summary>
    /// ラベル情報
    /// </summary>
    public string LabelName { get; }
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
        var irepo = configuration.GetSection("iRepoLink");

        DefTopName = irepo[$"{nameof(DefTopName)}"];
        SourceFolder = irepo[$"{nameof(SourceFolder)}"] ?? @".\Request";
        SourceFileName = irepo[$"{nameof(SourceFileName)}"] ?? "*.*";
        SuccessFileMoveForder = irepo[$"{nameof(SuccessFileMoveForder)}"] ?? @".\Success";
        ErrorFileMoveFolder = irepo[$"{nameof(ErrorFileMoveFolder)}"] ?? @".\Error";
        Encode = irepo[$"{nameof(Encode)}"] ?? "UTF-8";
        HeaderRowCount = int.TryParse(irepo[$"{nameof(HeaderRowCount)}"] ?? "0", out var rowCnt) ? rowCnt : 0;
        //DefTopId = int.TryParse(irepo[$"{nameof(DefTopId)}"] ?? , out var topId) ? topId : 0;
        if (irepo[$"{nameof(DefTopId)}"] is null) {
            DefTopId = null;
        } else {
            DefTopId = int.Parse(irepo[$"{nameof(DefTopId)}"] ?? "0");
        }
        RepName = irepo[$"{nameof(RepName)}"] ?? "{defTopName}_{datetime}";
        LabelName = irepo[$"{nameof(LabelName)}"] ?? "";
        SeparateChar = irepo[$"{nameof(SeparateChar)}"] ?? "AUTO";
        foreach (var x in irepo.GetSection($"{nameof(FromTo)}").GetChildren() ) {
            //            x.GetSection("ClusterID")
            var d = new FromTo();
            d.ColumnNo = int.Parse(x["ColumnNo"] ?? "-1");
            d.SheetNo = int.Parse(x["SheetNo"] ?? "-1");
            d.ClusterID = int.Parse(x["ClusterID"] ?? "-1");

            d.FromDateFormat = x["FromDateFormat"];
            d.ToDateFormat = x["ToDateFormat"];

            FromTo.Add(d);
        }
        //FromTo = irepo.GetSection($"{nameof(FromTo)}").Get<List<FromTo>>() ?? [];

    }
}
