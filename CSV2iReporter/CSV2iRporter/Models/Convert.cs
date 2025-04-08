//using ICSharpCode.SharpZipLib.Zip.Compression.Streams;
using Microsoft.Extensions.Configuration;
using Serilog;
using System.Text.Json.Serialization;
using System.Text;
using System.Runtime;
using System.ComponentModel.Design;

namespace CSV2iReporter.Models;


/// <summary>
/// 備考に割り当てるデータ
/// </summary>
public class Remark() {
    /// <summary>
    /// 変換元ファイルのカラム名または番号（列名または1～の名称）
    /// </summary>
    public int No { get; set; }

    /// <summary>
    /// 出力文字列
    /// </summary>
    public string Output { get; set; } = "";
}

/// <summary>
/// システムキー1～5に割り当てるデータ
/// </summary>
public class SystemKey() {

    /// <summary>
    /// 変換元ファイルのカラム名または番号（列名または1～の名称）
    /// </summary>
    public int No { get; set; }

    /// <summary>
    /// 出力文字列
    /// </summary>
    public string Output { get; set; } = "";

}

/// <summary>
/// クラスタに割り当てる固定
/// </summary>
public class Cluster() {

    /// <summary>
    /// 変換元ファイルのカラム名または番号（列名または1～の名称）
    /// </summary>
    public int Id { get; set; }

    /// <summary>
    /// 出力文字列
    /// </summary>
    public string Output { get; set; } = "";

}

public class Sheet(int no) {
    public int No { get; set; } = no;
    public List<Cluster> Cluster { get; set; } = [];
}

/// <summary>
/// CSVの列情報
/// </summary>
public class CsvColumn {
    /// <summary>
    /// 列番号（１～）
    /// </summary>
    public int No { get; set; }
    /// <summary>
    /// 列が日付の場合の書式
    /// </summary>
    public string? DateFormat { get; set; }
    /// <summary>
    /// 列が数値の場合の書式
    /// </summary>
    public string? NumericFormat { get; set; }
}

///// <summary>
///// 転送元の列からクラスターに割り当てるデータ
///// </summary>
//public class FromTo() {

//    /// <summary>
//    /// 変換元ファイルのカラム名または番号（列名または1～の名称）
//    /// </summary>
//    public int ColumnNo { get; set; }

//    /// <summary>
//    /// 変換先帳票のシートNo
//    /// </summary>
//    public int? SheetNo { get; set; }

//    /// <summary>
//    /// 変換先帳票のクラスターNo
//    /// </summary>
//    public int? ClusterID { get; set; }

//    /// <summary>
//    /// 帳票備考番号
//    /// </summary>
//    public int? RemarksNo { get; set; }

//    /// <summary>
//    /// システムキー番号
//    /// </summary>
//    public int? SystemKeyNo { get; set; }

//    /// <summary>
//    /// 日付型の場合の出力フォーマット
//    /// </summary>
//    public string? ToDateFormat { get; set; }

//    /// <summary>
//    /// 日付型の場合の入力フォーマット
//    /// </summary>
//    public string? FromDateFormat { get; set; }

//}


/// <summary>
/// 設定情報
/// </summary>
public class Convert() {

    /// <summary>
    /// 変換元ファイルが保存されたフォルダ名
    /// </summary>
    public string InputPath { get; set; } = @".\Input";
    /// <summary>
    /// <summary>
    /// 変換元ファイル名（ワイルドカード可）
    /// </summary>
    public string InputFilePattern { get; set; } = @"*.csv";
    /// 変換先ファイルが保存されるフォルダ名
    /// </summary>
    public string? OutputPath { get; set; } = null;
    /// 変換先ファイルの出力ファイル名
    /// </summary>
    public string? OutputFileName { get; set; } = @"{DefTopName}_{FileDateTime}.xml";
    /// <summary>
    /// 変換正常終了後のファイルの移動先フォルダ名
    /// </summary>
    public string SuccessFileMovePath { get; set; } = @".\Sucess";
    /// <summary>
    /// 変換異常終了後のファイルの移動先フォルダ名
    /// </summary>
    public string ErrorFileMovePath { get; set; } = @".\Error";
    /// <summary>
    /// データ転送元CSVファイルのエンコード
    /// </summary>
    public string Encode { get; set; } = @".\utf-8";
    /// <summary>
    /// ヘッダ行数
    /// </summary>
    public bool HasHeaderRecord { get; set; } = false;
    /// <summary>
    /// データ転送先の帳票定義ID
    /// </summary>
    public int? DefTopId { get; set; } = null;
    /// <summary>
    /// データ転送先の帳票定義名
    /// </summary>
    public string? DefTopName { get; set; } = null;
    /// <summary>
    /// 生成される帳票名称
    /// </summary>
    public string? RepTopName { get; set; } = null;
    /// <summary>
    /// ラベル情報
    /// </summary>
    public string? LabelName { get; set; } = null;
    /// <summary>
    /// CSV区切り文字
    /// </summary>
    public string Delimiter { get; set; } = @"auto";
    /// <summary>
    /// 備考1～10に設定されるデータ
    /// </summary>
    public Remark[]? Remarks { get; set; }
    /// <summary>
    /// システムキー1～5に設定されるデータ
    /// </summary>
    public SystemKey[]? SystemKeys { get; set; }
    /// <summary>
    /// クラスタに割り当てられるデータ
    /// </summary>
    public Sheet[]? Sheet { get; set; }
    /// <summary>
    /// CSVファイルの列フォーマット指定
    /// </summary>
    public CsvColumn[]? CsvColumn { get; set; }
}
