using AutoReportMake.Services;
using ConMasIReporterLib.Models;
using ExcelLib;
using Microsoft.Extensions.Configuration;
//using NPOI.OpenXmlFormats.Wordprocessing;
//using NPOI.SS.Formula;
//using Org.BouncyCastle.Ocsp;
using System.Xml.Linq;
using Serilog;
using System.Text;
using Org.BouncyCastle.Ocsp;
using Org.BouncyCastle.Utilities;
//using System.Text.RegularExpressions;

namespace CSV2iReporter;
/// <summary>
/// 帳票データ
/// </summary>
/// <remarks>
/// コンストラクタ
/// </remarks>
public class ReportData(int columnNo, int shrrtNo, int clusterId, string clusterData) {
    /// <summary>
    /// CSVの列番号
    /// </summary>
    public int columnNo = columnNo;
    /// <summary>
    /// シートNo
    /// </summary>
    public int sheetNo = shrrtNo;
    /// <summary>
    /// クラスターID
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
    public string _defTopName = "";
    /// <summary>
    /// 生成する帳票名
    /// </summary>
    public string _repTopName = "";
    /// <summary>備考情報1</summary>
    public string _remarksValue1 = "";
    /// <summary>備考情報2</summary>
    public string _remarksValue2 = "";
    /// <summary>備考情報3</summary>
    public string _remarksValue3 = "";
    /// <summary>備考情報4</summary>
    public string _remarksValue4 = "";
    /// <summary>備考情報5</summary>
    public string _remarksValue5 = "";
    /// <summary>備考情報6</summary>
    public string _remarksValue6 = "";
    /// <summary>備考情報7</summary>
    public string _remarksValue7 = "";
    /// <summary>備考情報8</summary>
    public string _remarksValue8 = "";
    /// <summary>備考情報9</summary>
    public string _remarksValue9 = "";
    /// <summary>備考情報10</summary>
    public string _remarksValue10 = "";
    /// <summary>クラスターID</summary>
    public List<int> _clusterId = [];

    /// <summary>
    /// 初期化
    /// </summary>
    public async Task<bool> Init() {
        if (!File.Exists($@"{AutoReportLib.Init.SettingFileFullPath}")) {
            Log.Error($"設定ファイル({AutoReportLib.Init.SettingFileFullPath})が ありません。");
            return false;
        }

        // 設定ファイルチェック
        if ( (_settings.DefTopId == 0) && (_settings.DefName == "") ) {
            Log.Error($"設定ファイル({AutoReportLib.Init.SettingFileFullPath})に DefTopId または DefName の設定がありません。");
            return false;
        }
        if (_settings.HeaderRowCount < 0) {
            Log.Error($"設定ファイル({AutoReportLib.Init.SettingFileFullPath})の HeaderRowCount の値が不正です。 値は 0～ を設定して下さい。");
            return false;
        }
        if (_settings.FromTo.Count == 0) {
            Log.Error($"設定ファイル({AutoReportLib.Init.SettingFileFullPath})に FromTo の設定がありません。");
            return false;
        } else {
            foreach(var ft in _settings.FromTo) {
                if (ft.SheetNo == 0) {
                    Log.Error($"設定ファイル({AutoReportLib.Init.SettingFileFullPath})の SheetNo の値が不正です。 値は 1～ を設定して下さい。");
                    return false;
                }
                if (ft.ColumnNo == 0) {
                    Log.Error($"設定ファイル({AutoReportLib.Init.SettingFileFullPath})の ColumnNo の値が不正です。 値は 1～ を設定して下さい。");
                    return false;
                }
                if (ft.ClusterID == -1) {
                    Log.Error($"設定ファイル({AutoReportLib.Init.SettingFileFullPath})の clusterID の値が不正です。 値は 0～ を設定して下さい。");
                    return false;
                }
            }
        }


        int defTopId = _settings.DefTopId;

        if(defTopId == 0) {

            // 名勝指定の場合

            // 指定の定義情報を読み込む
            var reports = await _svc.Req定義一覧取得Async(_settings.DefName);

            if (reports != null) {
                if (!reports.IsEmpty) {
                    // コレクション更新日の降順に帳票定義を並べ替えて最初の値を取り出す。nullの場合は0を設定する）
                    defTopId = (from item in reports.Elements("items").Elements("item") 
                                let updateDateTime = (DateTime?)item.Element("updateTime") 
                                let id = (int?)item.Element("itemId") 
                                orderby updateDateTime descending
                                select id).First() ?? 0;
                }
            }
            if (defTopId == 0) {
                Log.Error($"帳票定義名称({_settings.DefName})を含む帳票定義が見つかりません。");
                return false;
            }
        }


        // 帳票情報を取得
        var conmas = await _svc.Req定義簡易詳細情報取得Async(defTopId);
        if (conmas != null) {
            if (!conmas.IsEmpty) {
                var detailInfo = conmas.Element("detailInfo");

                // 帳票名
                _defTopName = detailInfo?.Element("topName")?.Value ?? "";
                // 備考情報
                _remarksValue1 = detailInfo?.Element("remarksValue1")?.Value ?? "";
                _remarksValue2 = detailInfo?.Element("remarksValue2")?.Value ?? "";
                _remarksValue3 = detailInfo?.Element("remarksValue3")?.Value ?? "";
                _remarksValue4 = detailInfo?.Element("remarksValue4")?.Value ?? "";
                _remarksValue5 = detailInfo?.Element("remarksValue5")?.Value ?? "";
                _remarksValue6 = detailInfo?.Element("remarksValue6")?.Value ?? "";
                _remarksValue7 = detailInfo?.Element("remarksValue7")?.Value ?? "";
                _remarksValue8 = detailInfo?.Element("remarksValue8")?.Value ?? "";
                _remarksValue9 = detailInfo?.Element("remarksValue9")?.Value ?? "";
                _remarksValue10 = detailInfo?.Element("remarksValue10")?.Value ?? "";
                // クラスター
                var clusters = detailInfo?.Element("clusters");
                var cluster = clusters?.Elements("cluster");
                if (cluster != null) {
                    foreach (var clusterelm in cluster) {
                        var name = clusterelm?.Element("name")?.Value;
                        var sheetNo = clusterelm?.Element("sheetNo")?.Value;
                        var clusterId = clusterelm?.Element("clusterId")?.Value;
                        if (int.TryParse(clusterId, out var id)) {
                            _clusterId.Add(id);
                        }
                    }
                }
            } else {
                 Log.Error($"帳票定義({_settings.DefTopId}) が読めませんでした。");
                return false;
            }
        } else {
            Log.Error($"ComMas WebAPI にログインできませんでした。");
            return false;
        }

        // 帳票定義にクラスターの有無チェック
        if (_clusterId.Count == 0) {
            Log.Error($"帳票定義({_settings.DefTopId}) にクラスター定義がありません。");
            return false;
        }

        // 設定ファイルのクラスターIDが帳票定義にあるかチェック
        foreach (var ft in _settings.FromTo) {
            if (!_clusterId.Contains(ft.ClusterID)) {
                Log.Error($"設定ファイル({AutoReportLib.Init.SettingFileFullPath})の clusterID({ft.ClusterID}) は 帳票定義({_settings.DefTopId}) にありません。");
                return false;
            }
        }

        // 監視、正常時、異常時データ格納先のディレクトリ生成
        if (!DirMake(_settings.SourceFileFolder)) {
            return false;
        }
        if (!DirMake(_settings.SuccessFileForder)) {
            return false;
        }
        if (!DirMake(_settings.ErrorFileFolder)) {
            return false;
        }

        return true;
    }

    /// <summary>
    /// ファイル情報を元に文字列変換を行う
    /// </summary>
    /// <param name="name"></param>
    /// <param name="fi"></param>
    /// <returns></returns>
    private static string ConvertNameByFile(string name, FileInfo fi) {
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
    private string ConvertNameByReport(string name) {
        // 帳票ID
        name = name.Replace("{defTopID}", _settings.DefTopId.ToString(), StringComparison.CurrentCultureIgnoreCase);
        // 帳票名
        name = name.Replace("{defTopName}", _defTopName, StringComparison.CurrentCultureIgnoreCase);
        // 備考
        name = name.Replace("{remarks1}", $"{_remarksValue1}", StringComparison.CurrentCultureIgnoreCase);
        name = name.Replace("{remarks2}", $"{_remarksValue2}", StringComparison.CurrentCultureIgnoreCase);
        name = name.Replace("{remarks3}", $"{_remarksValue3}", StringComparison.CurrentCultureIgnoreCase);
        name = name.Replace("{remarks4}", $"{_remarksValue4}", StringComparison.CurrentCultureIgnoreCase);
        name = name.Replace("{remarks5}", $"{_remarksValue5}", StringComparison.CurrentCultureIgnoreCase);
        name = name.Replace("{remarks6}", $"{_remarksValue6}", StringComparison.CurrentCultureIgnoreCase);
        name = name.Replace("{remarks7}", $"{_remarksValue7}", StringComparison.CurrentCultureIgnoreCase);
        name = name.Replace("{remarks8}", $"{_remarksValue8}", StringComparison.CurrentCultureIgnoreCase);
        name = name.Replace("{remarks9}", $"{_remarksValue9}", StringComparison.CurrentCultureIgnoreCase);
        name = name.Replace("{remarks10}", $"{_remarksValue10}", StringComparison.CurrentCultureIgnoreCase);
        return name;
    }
    /// <summary>
    /// データを元に文字列変換を行う
    /// </summary>
    /// <param name="name"></param>
    /// <param name="repData"></param>
    /// <param name="row"></param>
    /// <returns></returns>
    private static string ConvertNameByData(string name, List<ReportData> repData, int row) {

        // 行番号
        name = name.Replace("{rowNo}", row.ToString(), StringComparison.CurrentCultureIgnoreCase);
        // 列データの変換
        for (var col = 0; col < repData.Count; col++) {
            name = name.Replace("{" + repData[col].columnNo + "}", repData[col].clusterData, StringComparison.CurrentCultureIgnoreCase);
        }
        return name;
    }
    /// <summary>
    /// 生成帳票名の生成
    /// </summary>
    private string GetRepTopName(FileInfo fi, List<ReportData> repData, int row) {
        var name = ConvertNameByFile(_settings.RepTopName, fi);
        name = ConvertNameByReport(name);
        name = ConvertNameByData(name,repData,row);
        return name;
    }

    /// <summary>
    /// 帳票生成実行
    /// </summary>
    public async Task<bool> Execute() {
        // 監視ディレクトリ検索
        var files = Directory.GetFiles(_settings.SourceFileFolder, _settings.SourceFileName);

        // 見つかったファイル分の生成処理を行う
        foreach ( var file in files ) {
            var fi = new FileInfo(file);

            // 帳票の生成
            var ret = await MakeReport(fi);

            // 成功、失敗にによるファイルの移動
            if (ret == true) {
                try {
                    File.Move(fi.FullName, $"{_settings.SuccessFileForder}\\{fi.Name}", true);
                } catch (Exception ex) {
                    Log.Error(ex, $"変換元ファイル({fi.Name}) が {_settings.SuccessFileForder} に移動できませんでした。");
                    return false;
                }
            } else {
                try {
                    File.Move(fi.FullName, $"{_settings.ErrorFileFolder}\\{fi.Name}", true);
                } catch (Exception ex) {
                    Log.Error(ex, $"変換元ファイル({fi.Name}) が {_settings.ErrorFileFolder} に移動できませんでした。");
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
        var reportData = new List<ReportData>();
        var reqData = new List<string[]>();

        // ファイルの拡張子による変更元データ読込み
        if (fi.Extension.Equals(".csv", StringComparison.CurrentCultureIgnoreCase)) {
            reqData = GetRequestDataFromCsv(fi);
            if (reqData.Count == 0) {
                return false;
            }
        } else if (fi.Extension.Equals(".xlsx", StringComparison.CurrentCultureIgnoreCase)) {
            reqData = GetRequestDataFromExcel(fi);
            if (reqData.Count == 0) {
                return false;
            }
        } else {
            Log.Warning($"変換元ファイル({fi.Name}) の形式が違います。");
            return false;
        }

        // var remarks = _settings.Remarks;
        // 読み込んだデータから帳票を作成する
        for (var row = _settings.HeaderRowCount + 1; ; row++) {
            // 送り込むデータ情報をクリア
            reportData.Clear();

            //クラスタ情報がある場合
            var IsColumnNoExists = false;
            foreach (var clm in _settings.FromTo.Where(m => m.ClusterID != -1)) {
                // カラム情報
                string clmCell;
                // カラムデータを取出し
                clmCell = GetColumnData(reqData, row, clm.ColumnNo);

                if( clm.FromFormat is not null ) {
                    if( clm.ToFormat is not null ) {
                        clmCell = DateTime.ParseExact(clmCell, clm.FromFormat,null).ToString(clm.ToFormat);
                    } else {
                        Log.Error($"設定ファイル({AutoReportLib.Init.SettingFileFullPath})の FromTo の値が不正です。 FromFormatを指定した場合はToFormatを指定してください。");
                        return false;
                    }
                }
                if (!string.IsNullOrEmpty(clmCell)) {
                    IsColumnNoExists = true;
                }
                if (!string.IsNullOrEmpty(clmCell)) {
                    // 読み込んだ情報を設定情報に設定
                    reportData.Add(new ReportData((int)clm.ColumnNo, clm.SheetNo, clm.ClusterID, clmCell ));
                }
            }

            // 設定情報が１つも無かった場合、Excelの行読出し終了
            if (!IsColumnNoExists) break;

            // 生成する帳票名を生成する
            //_repTopName = GetRepTopName(fi, reportData, reqData[row], row - _settings.HeaderRowCount);
            var name = ConvertNameByFile(_settings.RepTopName, fi);
            name = ConvertNameByReport(name);
            name = ConvertNameByData(name, reportData, row);
            _repTopName = name;

            // ラベル情報設定
            var label = ConvertNameByFile(_settings.LabelName, fi);
            label = ConvertNameByReport(label);
            label = ConvertNameByData(label, reportData, row);


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
                .BeginR(_settings.DefTopId)
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
                Log.Error($"自動帳票作成要求に失敗しました。");
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// ディレクトリ作成
    /// </summary>
    private static bool DirMake(string path) {
        var dir = path.Replace("\\", "/");
        var net = dir[0..2] == "//" ? "//" : "";
        var dirs = dir.Split('/');

        var root = "";

        try {
            foreach (var itm in dirs) {
                if (string.IsNullOrEmpty(itm)) continue;
                if ((root != "") || (net == "")) {
                    if (!Directory.Exists($"{net}{root}{itm}")) {
                        Directory.CreateDirectory($"{net}{root}{itm}");
                    }
                }

                root += $"{itm}/";
            }
        } catch (Exception ex) {
            Log.Error(ex, $"ディレクトリ({path}) の作成に失敗しました。");
            return false;

        }

        return true;
    }

    /// <summary>
    /// 変換元ファイル読み込み（CSV）
    /// </summary>
    private List<string[]> GetRequestDataFromCsv(FileInfo fi) {
        var requestData = new List<string[]>();
        char splitChar;

        // CSVファイルのエンコード形式設定
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var encoding = _settings.CsvFileEncode == "SJIS" ? Encoding.GetEncoding("shift-jis") : Encoding.UTF8 ;

        try {
            // 全行読み込み
            var csv = File.ReadAllText(fi.FullName, encoding);

            // 区切り文字の判別
            switch (_settings.SplitType.ToUpper()) {
                case "TAB":
                case "\t":
                    splitChar = '\t';
                    break;
                case "COMMA":
                case ",":
                    splitChar = ',';
                    break;
                default:
                    // タブ文字が存在する場合はタブ区切りと判断
                    if (csv.IndexOf('\t') > 0) {
                        splitChar = '\t';
                    } else {
                        splitChar = ',';
                    }
                    break;

            }

            string[] lines;

            // 行単位に分割
            if ( csv.IndexOf("\r\n") > 0 ) {
                // CRLFで区切られている場合
                lines = csv.Split("\r\n");
            }else if (csv.IndexOf('r') > 0 ) {
                // CRのみで区切られている場合
                lines = csv.Split('\r');
            } else  {
                // LFのみで区切られている場合
                lines = csv.Split('\n');
            }

            // 桁単位に分割
            foreach (var line in lines) {
                requestData.Add($"{line}".Split(splitChar));
            }
        } catch (Exception ex) {
            Log.Error(ex, $"変換元ファイル({fi.Name}) の読込みに失敗しました。ファイルを開いている場合は閉じてください。");
            return [];
        }

        return requestData;
    }

    /// <summary>
    /// 変換元ファイル読み込み（Excel）
    /// </summary>
    private List<string[]> GetRequestDataFromExcel(FileInfo fi) {
        var svc = new NPOIExcelService();
        var requestData = new List<string[]>();

        try {
            // ワークブックオープン
            using var workbook = svc.Open(fi.FullName);

            // ワークシートを取出す
            var worksheet = workbook.GetWorkSheet(_settings.ExcelSheetName);
            // シート有りか？
            if ( !(worksheet.IsSheetAvailable) ) {
                Log.Error($"変換元ファイル({fi.Name}) に シート({_settings.ExcelSheetName}) がありません。");
                return [];
            }
            for (var row = 1; row <= worksheet.GetLastRowNum(); row++) {
                    
                // ラインデータ取り出し
                var rowCell = worksheet.GetRow(row);
                var clmData = new List<string>();
                    
                // ラインデータが無い場合は抜ける
                if (rowCell is null) {
                    continue;
                }

                //列データの取り出し
                for (var clm = 1; clm <= rowCell.GetLastCellNum(); clm++) {
                    // カラムデータ取り出し
                    var clmCell = rowCell.GetCell(clm);
                    // カラムデータ有りか？
                    if (clmCell != null) {
                        if( clmCell.CellType == CellType.String) {
                            clmData.Add(clmCell.StringCellValue);
                        } else {
                            clmData.Add(clmCell.GetStringValue());
                        }
                    } else {
                        clmData.Add("");
                    }
                }
                requestData.Add([.. clmData]);
            }
            
        } catch (Exception ex) {
            Log.Error(ex, $"変換元ファイル({fi.Name}) の読込みに失敗しました。ファイルを開いている場合は閉じてください。");
            return [];
        }

        return requestData;
    }

    /// <summary>
    /// 読み出した変換元の情報を取出す
    /// </summary>
    private static string GetColumnData(List<string[]> dat, int row, int clm) {
        var result = "";
        var row0 = row - 1;
        var clm0 = clm - 1;

        // 行、桁位置の最低値チェック
        if ((row0 >= 0) && (clm0 >= 0)) {
            // データに指定行があるか？
            if (dat.Count > row0) {
                // データに指定桁があるか？
                if (dat[row0].Length > clm0) {
                    result = dat[row0][clm0];
                }
            }
        }

        return result;
    }
}
