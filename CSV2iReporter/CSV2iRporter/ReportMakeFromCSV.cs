﻿using AutoReportMake.Services;
using ConMasIReporterLib.Models;
//using ExcelLib;
using Microsoft.Extensions.Configuration;
using System.Xml.Linq;
using Serilog;
using System.Text;
using System.Globalization;

namespace CSV2iReporter;
/// <summary>
/// 帳票データ
/// </summary>
/// <remarks>
/// コンストラクタ
/// </remarks>
public class ReportData(int columnNo, int sheetNo, int clusterId, string clusterData) {
    /// <summary>
    /// データの列番号（1～）
    /// </summary>
    public int columnNo = columnNo;
    /// <summary>
    /// シートNo（1～）
    /// </summary>
    public int sheetNo = sheetNo;
    /// <summary>
    /// クラスターID（0～）
    /// </summary>
    public int clusterId = clusterId;
    /// <summary>
    /// クラスターデータ
    /// </summary>
    public string clusterData = clusterData;
}

/// <summary>
/// i-Reporter自動帳票連携ソフト
/// </summary>
/// <remarks>
/// コンストラクタ
/// </remarks>
public class ReportMakeFromCSV(IConfigurationRoot configuration) {
    private readonly Settings _settings = new(configuration);
    private readonly ComMasWebAPIService2 _svc = new(configuration.GetSection("ConMasWeb"));

    /// <summary>
    /// 帳票定義の帳票名
    /// </summary>
    private string? _defTopName;
    /// <summary>
    /// 処理対象の帳票定義ID
    /// </summary>
    private int? _defTopId;
    /// <summary>
    /// 生成する帳票名
    /// </summary>
    private string? _repTopName;
    ///// <summary>備考情報1</summary>
    //private string _remarksValue1 = "";
    ///// <summary>備考情報2</summary>
    //private string _remarksValue2 = "";
    ///// <summary>備考情報3</summary>
    //private string _remarksValue3 = "";
    ///// <summary>備考情報4</summary>
    //private string _remarksValue4 = "";
    ///// <summary>備考情報5</summary>
    //private string _remarksValue5 = "";
    ///// <summary>備考情報6</summary>
    //private string _remarksValue6 = "";
    ///// <summary>備考情報7</summary>
    //private string _remarksValue7 = "";
    ///// <summary>備考情報8</summary>
    //private string _remarksValue8 = "";
    ///// <summary>備考情報9</summary>
    //private string _remarksValue9 = "";
    ///// <summary>備考情報10</summary>
    //private string _remarksValue10 = "";
    ///// <summary>
    ///// クラスターIDリスト
    ///// </summary>
    //private readonly List<int> _clusterId = [];
    /// <summary>
    /// 列名称ヘッダ配列
    /// </summary>
    private string[]? _columnHeaders = null;


    /// <summary>
    /// 初期化
    /// </summary>
    public async Task<bool> Init() {
        if (!File.Exists($@"{AutoReportLib.Init.SettingFileFullPath}")) {
            Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'が ありません");
            return false;
        }

        // 設定ファイルチェック
        if ( (_settings.DefTopId is null) && (_settings.DefTopName is null) ) {
            Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'に DefTopId または DefTopName の設定がありません");
            return false;
        }
        if (_settings.HeaderRowCount < 0) {
            Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'の HeaderRowCount の値が不正です。 値は 0～ を設定して下さい");
            return false;
        }
        if (_settings.FromTo.Count == 0) {
            Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'に FromTo の設定がありません");
            return false;
        } else {
            foreach(var ft in _settings.FromTo) {
                //if (ft.Column == "") {
                //    Log.Error($"設定ファイル({AutoReportLib.Init.SettingFileFullPath})の Column の値が不正です。 値は 1～ を設定して下さい");
                //    return false;
                //}
                if (ft.SheetNo == 0) {
                    Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'の SheetNo の値が不正です。 値は 1～ を設定して下さい");
                    return false;
                }
                if (ft.ClusterID == -1) {
                    Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'の clusterID の値が不正です。 値は 0～ を設定して下さい");
                    return false;
                }
            }
        }


        _defTopName = _settings.DefTopName;
        _defTopId = _settings.DefTopId;

        if(_defTopId is null) {

            if(_defTopName is null) {
                Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'に DefTopId または DefTopName の設定がありません");
                return false;
            }

            // 名称指定の場合、最新の帳票IDを取得する

            // 指定の定義情報を読み込む
            var reports = await _svc.Req定義一覧取得Async(_defTopName);

            if (reports != null) {
                if (!reports.IsEmpty) {
                    // コレクション更新日の降順に帳票定義を並べ替えて最初の値を取り出す）
                    _defTopId = (from item in reports.Elements("items").Elements("item") 
                                let updateDateTime = (DateTime?)item.Element("updateTime") 
                                let id = (int?)item.Element("itemId") 
                                orderby updateDateTime descending
                                select id).First();
                }
            }
            if (_defTopId is null) {
                Log.Error($"帳票定義名称'{_defTopName}'を含む帳票定義が見つかりません");
                return false;
            }
        }


        // 帳票情報を取得
        var conmas = await _svc.Req定義簡易詳細情報取得Async(_defTopId??0);
        if (conmas != null) {
            if (!conmas.IsEmpty) {
                //var detailInfo = conmas.Element("detailInfo");

                //// 帳票名
                //_defTopName = detailInfo?.Element("topName")?.Value ?? "";
                ////// 備考情報
                ////_remarksValue1 = detailInfo?.Element("remarksValue1")?.Value ?? "";
                ////_remarksValue2 = detailInfo?.Element("remarksValue2")?.Value ?? "";
                ////_remarksValue3 = detailInfo?.Element("remarksValue3")?.Value ?? "";
                ////_remarksValue4 = detailInfo?.Element("remarksValue4")?.Value ?? "";
                ////_remarksValue5 = detailInfo?.Element("remarksValue5")?.Value ?? "";
                ////_remarksValue6 = detailInfo?.Element("remarksValue6")?.Value ?? "";
                ////_remarksValue7 = detailInfo?.Element("remarksValue7")?.Value ?? "";
                ////_remarksValue8 = detailInfo?.Element("remarksValue8")?.Value ?? "";
                ////_remarksValue9 = detailInfo?.Element("remarksValue9")?.Value ?? "";
                ////_remarksValue10 = detailInfo?.Element("remarksValue10")?.Value ?? "";
                //// クラスター
                //var clusters = detailInfo?.Element("clusters");
                //var cluster = clusters?.Elements("cluster");
                //if (cluster != null) {
                //    foreach (var clusterelm in cluster) {
                //        var name = clusterelm?.Element("name")?.Value;
                //        var sheetNo = clusterelm?.Element("sheetNo")?.Value;
                //        var clusterId = clusterelm?.Element("clusterId")?.Value;
                //        if (int.TryParse(clusterId, out var id)) {
                //            _clusterId.Add(id);
                //        }
                //    }
                //}
            } else {
                 Log.Error($"帳票定義'{_defTopId}' が読込できませんでした");
                return false;
            }
        } else {
            Log.Error($"ComMas WebAPI にログインできませんでした");
            return false;
        }

        //// 帳票定義にクラスターの有無チェック
        //if (_clusterId.Count == 0) {
        //    Log.Error($"帳票定義'{_defTopId}' にクラスター定義がありません");
        //    return false;
        //}

        // 設定ファイルのクラスターIDが帳票定義にあるかチェック
        foreach (var setting in _settings.FromTo) {
            var detailInfo = conmas.Element("detailInfo");

            // クラスター
            var clusters = detailInfo?.Element("clusters");
            var cluster = clusters?.Elements("cluster");
            string? sheetNo;
            string? clusterId;
            if (cluster != null) {
                // 定義されたシートNoとクラスターIDが存在するか？
                bool isFind = false;
                foreach (var clusterelm in cluster) {
                    sheetNo = clusterelm?.Element("sheetNo")?.Value;
                    clusterId = clusterelm?.Element("clusterId")?.Value;
                    if (int.TryParse(sheetNo, out var sheet)) {
                        if (int.TryParse(clusterId, out var id)) {
                            if (setting.SheetNo.Equals(sheet) && setting.ClusterID.Equals(id)) {
                                isFind = true;
                                break;
                            }
                        }
                    }
                }
                if (!isFind) {
                    Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'の sheetNo'{setting.SheetNo}'、clusterID'{setting.ClusterID}' は 帳票定義'{_defTopId}' にありません");
                    return false;
                }
            }
        }

        // 監視、正常時、異常時データ格納先のディレクトリ生成
        if( !Directory.Exists(_settings.SourceFolder)) {
            try {
                Directory.CreateDirectory(_settings.SourceFolder);
            } catch (Exception ex) {
                Log.Error(ex,$"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'の SourceFolder'{_settings.SourceFolder}'を作成できませんでした");
            }
        }
        if (!Directory.Exists(_settings.SuccessFileMoveForder)) {
            try {
                Directory.CreateDirectory(_settings.SuccessFileMoveForder);
            } catch (Exception ex) {
                Log.Error(ex,$"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'の SuccessFileMoveForder'{_settings.SuccessFileMoveForder}'を作成できませんでした");
            }
        }
        if (!Directory.Exists(_settings.ErrorFileMoveFolder)) {
            try {
                Directory.CreateDirectory(_settings.ErrorFileMoveFolder);
            } catch (Exception ex) {
                Log.Error(ex,$"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'の ErrorFileMoveFolder'{_settings.ErrorFileMoveFolder}'を作成できませんでした");
            }
        }

        //if (!DirMake(_settings.SourceFolder)) {
        //    return false;
        //}
        //if (!DirMake(_settings.SuccessFileMoveForder)) {
        //    return false;
        //}
        //if (!DirMake(_settings.ErrorFileMoveFolder)) {
        //    return false;
        //}

        return true;
    }

    /// <summary>
    /// ファイル情報を元に文字列変換を行う
    /// </summary>
    /// <param name="name"></param>
    /// <param name="fi"></param>
    /// <returns></returns>
    private static string ConvertNameFromFile(string name, FileInfo fi) {
        var dt = DateTime.Now;

        // 元ファイル名（ベース）
        name = name.Replace("{fileName}", Path.GetFileNameWithoutExtension(fi.Name), StringComparison.CurrentCultureIgnoreCase);
        // 依頼ファイルのタイムスタンプ（更新日時）
        name = name.Replace("{fileDate}", $"{fi.LastWriteTime:yyMMdd}", StringComparison.CurrentCultureIgnoreCase);
        name = name.Replace("{fileDateTime}", $"{fi.LastWriteTime::yyMMddHHmmss}", StringComparison.CurrentCultureIgnoreCase);
        name = name.Replace("{fileTime}", $"{fi.LastWriteTime::HHmmss}", StringComparison.CurrentCultureIgnoreCase);
        // 処理時の日時
        name = name.Replace("{date}", $"{dt:yyMMdd}", StringComparison.CurrentCultureIgnoreCase);
        name = name.Replace("{time}", $"{dt:HHmmss}", StringComparison.CurrentCultureIgnoreCase);
        name = name.Replace("{dateTime}", $"{dt:yyMMddHHmmss}", StringComparison.CurrentCultureIgnoreCase);
        return name;
    }
    /// <summary>
    /// 帳票定義を元に文字列変換を行う
    /// </summary>
    /// <param name="name"></param>
    /// <returns></returns>
    private string ConvertNameFromReport(string name) {
        // 帳票ID
        name = name.Replace("{defTopID}", _defTopId.ToString(), StringComparison.CurrentCultureIgnoreCase);
        // 帳票名
        name = name.Replace("{defTopName}", _defTopName, StringComparison.CurrentCultureIgnoreCase);
        //// 備考
        //name = name.Replace("{remarks1}", $"{_remarksValue1}", StringComparison.CurrentCultureIgnoreCase);
        //name = name.Replace("{remarks2}", $"{_remarksValue2}", StringComparison.CurrentCultureIgnoreCase);
        //name = name.Replace("{remarks3}", $"{_remarksValue3}", StringComparison.CurrentCultureIgnoreCase);
        //name = name.Replace("{remarks4}", $"{_remarksValue4}", StringComparison.CurrentCultureIgnoreCase);
        //name = name.Replace("{remarks5}", $"{_remarksValue5}", StringComparison.CurrentCultureIgnoreCase);
        //name = name.Replace("{remarks6}", $"{_remarksValue6}", StringComparison.CurrentCultureIgnoreCase);
        //name = name.Replace("{remarks7}", $"{_remarksValue7}", StringComparison.CurrentCultureIgnoreCase);
        //name = name.Replace("{remarks8}", $"{_remarksValue8}", StringComparison.CurrentCultureIgnoreCase);
        //name = name.Replace("{remarks9}", $"{_remarksValue9}", StringComparison.CurrentCultureIgnoreCase);
        //name = name.Replace("{remarks10}", $"{_remarksValue10}", StringComparison.CurrentCultureIgnoreCase);
        return name;
    }
    /// <summary>
    /// データを元に文字列変換を行う
    /// </summary>
    /// <param name="name"></param>
    /// <param name="repData"></param>
    /// <param name="row"></param>
    /// <returns></returns>
    private static string ConvertNameFromData(string name, List<ReportData> repData, int rowNo) {

        // 行番号
        name = name.Replace("{RowNo}", rowNo.ToString(), StringComparison.CurrentCultureIgnoreCase);
        // 列データの変換
        for (var col = 0; col < repData.Count; col++) {
            name = name.Replace("{" + repData[col].columnNo + "}", repData[col].clusterData, StringComparison.CurrentCultureIgnoreCase);
        }
        return name;
    }
    /// <summary>
    /// 文字列から取得する列番号を取得する
    /// </summary>
    /// <param name="columns">列名定義配列</param>
    /// <param name="columnName">探す列名</param>
    /// <returns>0～:列番号、null:列指定エラー</returns>
    private int? GetColumnNo(string columnName) {
        int colNo;

        
        if(this._columnHeaders is not null) {
            //列名を示す文字列から探す
            for (colNo = 0; colNo < this._columnHeaders.Length; colNo++) {
                // カラム名と設定文字列を比較する
                if (string.Compare(this._columnHeaders[colNo], columnName, StringComparison.OrdinalIgnoreCase) == 0) {
                    return colNo;
                }
            }
        }
        //列名と一致しない場合は、列番号として扱う
        if( !(int.TryParse(columnName,out colNo)) ) {
            return null;
        }
        return colNo;
    }
    /// <summary>
    /// 生成帳票名の生成
    /// </summary>
    private string GetRepTopName(FileInfo fi, List<ReportData> repData, int rowNo) {
        var name = ConvertNameFromFile(_settings.RepName, fi);
        name = ConvertNameFromReport(name);
        name = ConvertNameFromData(name,repData,rowNo);
        return name;
    }

    /// <summary>
    /// 帳票生成実行
    /// </summary>
    public async Task<bool> Execute() {
        // 監視ディレクトリ検索
        var files = Directory.GetFiles(_settings.SourceFolder, _settings.SourceFileName);

        // 見つかったファイル分の生成処理を行う
        foreach ( var file in files ) {
            var fi = new FileInfo(file);

            // 帳票の生成
            var ret = await MakeReport(fi);

            // 成功、失敗にによるファイルの移動
            if (ret == true) {
                try {
                    File.Move(fi.FullName, $"{_settings.SuccessFileMoveForder}\\{fi.Name}", true);
                } catch (Exception ex) {
                    Log.Error(ex, $"変換元ファイル'{fi.Name}' を処理正常フォルダ '{_settings.SuccessFileMoveForder}' に移動できませんでした");
                    return false;
                }
            } else {
                try {
                    File.Move(fi.FullName, $"{_settings.ErrorFileMoveFolder}\\{fi.Name}", true);
                } catch (Exception ex) {
                    Log.Error(ex, $"変換元ファイル'{fi.Name}' を処理失敗フォルダ '{_settings.ErrorFileMoveFolder}' に移動できませんでした");
                }
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// 帳票の生成を行う
    /// </summary>
    private async Task<bool> MakeReport(FileInfo fi) {
        // 元ファイルの全行のデータ
        var reqData = new List<string[]>();
        // 変換リスト
        var reportData = new List<ReportData>();

        // ファイルの拡張子による変更元データ読込み
        if (fi.Extension.Equals(".csv", StringComparison.CurrentCultureIgnoreCase)) {
            reqData = GetRequestDataFromCsv(fi);
            if (reqData.Count == 0) {
                return false;
            }
        } else {
            Log.Warning($"変換元ファイル'{fi.Name}' の形式が違います");
            return false;
        }

        // ヘッダ行を取出し
        if (_settings.HeaderRowCount > 0) {
            this._columnHeaders = reqData[_settings.HeaderRowCount - 1];
        }

        // 読み込んだデータから帳票を作成する
        for (var rowNo = _settings.HeaderRowCount; ; rowNo++) {

            // 送り込むデータ情報をクリア
            reportData.Clear();

            //クラスタ情報があるかを確認するフラグ
            bool IsDataExists = false;

            foreach (var fromTo in _settings.FromTo.Where(m => m.ClusterID != -1)) {

                // 列番号0～
                int? colNo;

                // カラム情報
                string? cellValue;

                // データ取得対象の列番号を取得
                // colNo = GetColumnNo(fromTo.ColumnNo);
                colNo = fromTo.ColumnNo;

                // 列番号が取得できなければエラーとする
                if (colNo is null) {
                    Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'の列番号'{colNo}'が見つかりません");
                    return false;
                }
                colNo = fromTo.ColumnNo;

                // カラムデータを取出し
                cellValue = GetColumnData(reqData, rowNo, (int)colNo);

                // カラムデータがない場合はエラー
                if (cellValue is null) {
                    Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'の列番号'{colNo}'列のデータがありません");
                    return false;
                }

                // 日付変換
                CultureInfo culture = new CultureInfo("ja-JP");

                // 日付文字列を変換して登録する
                // FromDateFormat
                if (fromTo.FromDateFormat is not null) {
                    DateTime dt;
                    // 日付文字列を変換して登録する
                    try {
                        dt = DateTime.ParseExact(cellValue, fromTo.FromDateFormat, null);
                    } catch (Exception ex) {
                        Log.Error(ex,$"CSVファイルの{colNo}列目の日付文字列'{cellValue}'または 設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'の FromDateFormat の指定'{fromTo.FromDateFormat}'が不正です");
                        return false;
                    }
                    //ToDateFormat
                    try {
                        if (fromTo.ToDateFormat is not null) {
                            cellValue = dt.ToString(fromTo.ToDateFormat);
                        } else {
                            cellValue = dt.ToString("yyyy/MM/dd HH:mm:ss");
                        }
                    } catch (Exception ex) {
                        Log.Error(ex,$"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'の ToDateFormat の指定'{fromTo.ToDateFormat}'が不正です");
                    }
                }


                // データが空でない場合データアリフラグをtrueに設定する
                if (!string.IsNullOrEmpty(cellValue)) {
                    IsDataExists = true;
                }

                // データがあれば追加する
                if (string.IsNullOrEmpty(cellValue)) {
                    reportData.Add(new ReportData((int)colNo, fromTo.SheetNo, fromTo.ClusterID, cellValue ));
                }
            }

            // 設定情報が１つも無かった場合、行読出し終了
            if (!IsDataExists) break;

            // 生成する帳票名を生成する
            //_repTopName = GetRepTopName(fi, reportData, reqData[row], row - _settings.HeaderRowCount);
            var name = ConvertNameFromFile(_settings.RepName, fi);
            name = ConvertNameFromReport(name);
            name = ConvertNameFromData(name, reportData, rowNo);
            _repTopName = name;

            // ラベル情報設定
            var label = ConvertNameFromFile(_settings.LabelName, fi);
            label = ConvertNameFromReport(label);
            label = ConvertNameFromData(label, reportData, rowNo);


            // 自動帳票作成を要求するCSVデータを作成する
            //  ヘッダ部
            var sb = CSVレイアウト
                // defTopId
                .BeginH()
                // repTopName
                .AddTopName();

            //// remarksValue
            //if (remarks != null) {
            //    foreach (var r in remarks.Remarks.OrderBy(m => m.Index)) {
            //        sb.AddTOP備考情報(r.Index);
            //    }
            //}
            // addLabels
            sb.addLabels();
            foreach (ReportData r in reportData) {
                // SxxCxx
                sb.Addクラスター(r.sheetNo, r.clusterId);
            }
            sb.End();

            //  データ部
            sb
                // defTopId
                .BeginR(_defTopId??0)
                // repTopName
                .AddValue(_repTopName);

            //// remarksValue
            //if (remarks != null) {
            //    foreach (var r in remarks.Remarks.OrderBy(m => m.Index)) {
            //        sb.AddValue(r.Value ?? "");
            //    }
            //}

            // addLabels
            sb.AddValue(label);
                ;
            foreach (var dat in reportData) {
                // SxxCxx
                sb.AddValue(dat.clusterData);
            }
            sb.End();

            // ConMasへ自動帳票作成要求を行う
            var reqresult = await _svc.Req自動帳票作成Csv2Async(sb.ToString());

            // 要求が失敗した場合は中止する
            if (!reqresult.IsSuccess) {
                Log.Error($"自動帳票作成要求に失敗しました。エラーコードは({reqresult.ErrorCode})です");
                return false;
            }
        }

        return true;
    }

    ///// <summary>
    ///// ディレクトリ作成
    ///// </summary>
    //private static bool DirMake(string path) {
    //    var dir = path.Replace("\\", "/");
    //    var net = dir[0..2] == "//" ? "//" : "";
    //    var dirs = dir.Split('/');

    //    var root = "";

    //    try {
    //        foreach (var itm in dirs) {
    //            if (string.IsNullOrEmpty(itm)) continue;
    //            if ((root != "") || (net == "")) {
    //                if (!Directory.Exists($"{net}{root}{itm}")) {
    //                    Directory.CreateDirectory($"{net}{root}{itm}");
    //                }
    //            }

    //            root += $"{itm}/";
    //        }
    //    } catch (Exception ex) {
    //        Log.Error(ex, $"転送先フォルダ'{path}' の作成に失敗しました");
    //        return false;

    //    }

    //    return true;
    //}

    /// <summary>
    /// 変換元ファイル読み込み（CSV）
    /// </summary>
    private List<string[]> GetRequestDataFromCsv(FileInfo fi) {
        var requestData = new List<string[]>();
        char separateChar;

        // CSVファイルのエンコード形式設定
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        // システムデフォルトのエンコード形式
        Encoding encoding = Encoding.Default;

        switch ( _settings.Encode.ToLower() )
        {
            case "sjis":
            case "shift-jis":
            case "shift_jis":
                encoding = Encoding.GetEncoding("shift-jis");
                break;
            case "utf8":
            case "utf-8":
            case "utf_8":
                encoding = Encoding.UTF8;
                break;
            case "unicode":
            case "utf-16":
            case "utf_16":
                encoding = Encoding.Unicode;
                break;
            case "euc-jp":
            case "euc_jp":
                encoding = Encoding.GetEncoding("euc-jp");
                break;
            case "iso-2022-jp":
            case "iso_2022_jp":
            case "jis":
                encoding = Encoding.GetEncoding("iso-2022-jp");
                break;
        }

        try {
            // 全行読み込み
            var csv = File.ReadAllText(fi.FullName, encoding);

            // 区切り文字の判別
            switch (_settings.SeparateChar.ToLower()) {
                case "tab":
                case "\t":
                case "\\t":
                    separateChar = '\t';
                    break;
                case "comma":
                case ",":
                    separateChar = ',';
                    break;
                default:
                    // タブ文字が存在する場合はタブ区切りと判断
                    if (csv.IndexOf('\t') > 0) {
                        separateChar = '\t';
                    } else {
                        separateChar = ',';
                    }
                    break;
            }

            string[] lines;

            // 行単位に分割
            if ( csv.IndexOf("\r\n") > 0 ) {
                // CRLFで区切られている場合
                lines = csv.Split("\r\n");
            }else if (csv.IndexOf('\r') > 0 ) {
                // CRのみで区切られている場合
                lines = csv.Split('\r');
            } else  {
                // LFのみで区切られている場合
                lines = csv.Split('\n');
            }

            // 桁単位に分割
            foreach (var line in lines) {
                requestData.Add($"{line}".Split(separateChar));
            }
        } catch (Exception ex) {
            Log.Error(ex, $"変換元ファイル'{fi.Name}' の読込みに失敗しました。ファイルを開いている可能性があります");
            return [];
        }

        return requestData;
    }

    /// <summary>
    /// 読み出した変換元の情報を取出す
    /// </summary>
    private static string? GetColumnData(List<string[]> dat, int row, int col) {
        // 行、桁位置の最低値チェック
        if ((row >= 0) && (col >= 0)) {
            // データに指定行があるか？
            if (dat.Count > row) {
                // データに指定桁があるか？
                if (dat[row].Length > col) {
                    return dat[row][col];
                }
            }
        }
        return null;
    }
}