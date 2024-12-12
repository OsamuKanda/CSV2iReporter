using AutoReportMake.Services;
using ConMasIReporterLib.Models;
//using ExcelLib;
using Microsoft.Extensions.Configuration;
using System.Xml.Linq;
using Serilog;
using System.Text;
using System.Globalization;
using System.Collections.Generic;
using System.Runtime;
using CsvHelper;
using System.Reflection.Metadata;
using CsvHelper.Configuration;
using System.Linq;

namespace CSV2iReporter;
/// <summary>
/// 帳票データ
/// </summary>
/// <remarks>
/// コンストラクタ
/// </remarks>
public class ReportData(FromTo fromTo, string data) {
    public FromTo fromto = fromTo;
    /// <summary>
    /// クラスターデータ
    /// </summary>
    public string data = data;
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
    ///// 列名称ヘッダ配列
    ///// </summary>
    //private string[]? _columnHeaders = null;


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
        // FromTo設定
        if (_settings.FromTo.Count == 0) {
            Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'に FromTo の記述が必須です");
            return false;
        } else {
            foreach(var ft in _settings.FromTo) {
                // 列番号
                if (ft.ColumnNo == -1) {
                    Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'の FromTo には ColumnNo の記述が必須です");
                    return false;
                }

                // 転送先
                if ((ft.SheetNo is null) && (ft.ClusterID is not null)) {
                    Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'の FromTo で ClusterID を指定した場合は SheetNo の記述が必須です");
                    return false;
                } else if ((ft.SheetNo is not null) && (ft.ClusterID is null)) {
                    Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'の FromTo で SheetNo を指定した場合は ClusterID の記述が必須です");
                    return false;
                }
                
                // シート番号確認
                if (ft.SheetNo < 1) {
                    Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'の SheetNo の値が不正です。 SheetNo には1以上の値を設定して下さい");
                    return false;
                }
                
                // クラスターID
                if (ft.ClusterID < 0) {
                    Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'の clusterID の値が不正です。 clusterID には0以上の値を設定して下さい");
                    return false;
                }
            }
        }


        _defTopName = _settings.DefTopName;
        _defTopId = _settings.DefTopId;

        if(_defTopId is null) {

            // 名称指定の場合、最新の帳票IDを取得する

            // 指定の定義情報を読み込む
            var reports = await _svc.Req定義一覧取得Async(_defTopName??"");

            if ((reports != null) && (!reports.IsEmpty)) {
                // 帳票定義名が一致する帳票定義から最新の更新日時の帳票定義IDを取得する
                // （更新日の降順に帳票定義を並べ替えて最初の帳票IDを取得）
                _defTopId = (from item in reports.Elements("items").Elements("item") 
                                let updateDateTime = (DateTime?)item.Element("updateTime") 
                                let id = (int?)item.Element("itemId") 
                                orderby updateDateTime descending
                                select id).First();
            }
            if (_defTopId is null) {
                Log.Error($"帳票定義名称が'{_defTopName}'の帳票定義が見つかりません");
                return false;
            }
        }


        // 帳票情報を取得
        var conmas = await _svc.Req定義簡易詳細情報取得Async(_defTopId??0);
        if (conmas is null) {
            Log.Error($"ComMas WebAPI にログインできませんでした");
            return false;
        }
        if (conmas.IsEmpty) {
            Log.Error($"帳票定義'{_defTopId}' が読込できませんでした");
            return false;
        }

        //// 帳票定義にクラスターの有無チェック
        //if (_clusterId.Count == 0) {
        //    Log.Error($"帳票定義'{_defTopId}' にクラスター定義がありません");
        //    return false;
        //}

        // 設定ファイルのクラスターIDが帳票定義にあるかチェック
        foreach (var fromTo in _settings.FromTo.Where(m => m.ClusterID is not null)) {
            var detailInfo = conmas.Element("detailInfo");

            // クラスター
            var clusters = detailInfo?.Element("clusters");
            var cluster = clusters?.Elements("cluster");
            string? sheetNo;
            string? clusterId;
            if (cluster != null) {
                // 定義されたシートNoとクラスターIDが存在するか？
                bool isFind = false;
                foreach (var repCluster in cluster) {

                    sheetNo = repCluster?.Element("sheetNo")?.Value;
                    clusterId = repCluster?.Element("clusterId")?.Value;

                    if ( !int.TryParse(sheetNo, out var sheet) )
                        continue;

                    if ( !int.TryParse(clusterId, out var id) )
                        continue;

                    if ( (!fromTo.SheetNo.Equals(sheet)) || (!fromTo.ClusterID.Equals(id)) )
                        continue;

                    //定義されたシートNoとクラスターIDが存在した
                    isFind = true;
                    break;
                }
                if (!isFind) {
                    Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'の sheetNo'{fromTo.SheetNo}'、clusterID'{fromTo.ClusterID}' は 帳票定義'{_defTopId}' にありません");
                    return false;
                }
            }
        }

        // 監視、正常時、異常時データ格納先のディレクトリが無ければ生成
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

        return true;
    }

    /// <summary>
    /// ファイル情報を元に文字列変換を行う
    /// </summary>
    /// <param name="repTopName"></param>
    /// <param name="fi"></param>
    /// <returns></returns>
    private static string ConvertNameFromFile(string repTopName, FileInfo fi) {
        var dt = DateTime.Now;

        // 元ファイル名（ベース）
        repTopName = repTopName.Replace("{fileName}", Path.GetFileNameWithoutExtension(fi.Name), StringComparison.CurrentCultureIgnoreCase);
        // 依頼ファイルのタイムスタンプ（更新日時）
        repTopName = repTopName.Replace("{fileDate}", $"{fi.LastWriteTime:yyMMdd}", StringComparison.CurrentCultureIgnoreCase);
        repTopName = repTopName.Replace("{fileDateTime}", $"{fi.LastWriteTime::yyMMddHHmmss}", StringComparison.CurrentCultureIgnoreCase);
        repTopName = repTopName.Replace("{fileTime}", $"{fi.LastWriteTime::HHmmss}", StringComparison.CurrentCultureIgnoreCase);
        // 処理時の日時
        repTopName = repTopName.Replace("{date}", $"{dt:yyMMdd}", StringComparison.CurrentCultureIgnoreCase);
        repTopName = repTopName.Replace("{time}", $"{dt:HHmmss}", StringComparison.CurrentCultureIgnoreCase);
        repTopName = repTopName.Replace("{dateTime}", $"{dt:yyMMddHHmmss}", StringComparison.CurrentCultureIgnoreCase);
        return repTopName;
    }
    /// <summary>
    /// 帳票定義を元に文字列変換を行う
    /// </summary>
    /// <param name="repTopName"></param>
    /// <returns></returns>
    private string ConvertNameFromReport(string repTopName) {
        // 帳票ID
        repTopName = repTopName.Replace("{defTopID}", _defTopId.ToString(), StringComparison.CurrentCultureIgnoreCase);
        // 帳票名
        repTopName = repTopName.Replace("{defTopName}", _defTopName, StringComparison.CurrentCultureIgnoreCase);

        return repTopName;
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
            name = name.Replace("{" + repData[col].fromto.ColumnNo + "}", repData[col].data, StringComparison.CurrentCultureIgnoreCase);
        }
        return name;
    }
    /// <summary>
    /// 帳票生成実行
    /// </summary>
    public async Task<bool> Execute() {
        // 監視ディレクトリ検索
        var files = Directory.GetFiles(_settings.SourceFolder, _settings.SourceFileName).OrderBy(f => File.GetLastWriteTime(f));

        // 見つかったファイル分の生成処理を行う
        foreach ( var file in files ) {
            var fi = new FileInfo(file);

            Log.Information($"ファイル'{fi.FullName}'処理開始");

            // 帳票の生成
            var ret = await MakeReport(fi);

            // 成功、失敗にによるファイルの移動
            if (ret == true) {
                Log.Information($"正常処理ファイル'{fi.FullName}'移動先'{_settings.SuccessFileMoveForder}\\{fi.Name}'");
                try {
                    File.Move(fi.FullName, $"{_settings.SuccessFileMoveForder}\\{fi.Name}", true);
                } catch (Exception ex) {
                    Log.Error(ex, $"変換元ファイル'{fi.Name}' を処理正常フォルダ '{_settings.SuccessFileMoveForder}' に移動できませんでした");
                    return false;
                }
            } else {
                Log.Information($"異常処理ファイル'{fi.FullName}'移動先'{_settings.ErrorFileMoveFolder}\\{fi.Name}'");
                try {
                    File.Move(fi.FullName, $"{_settings.ErrorFileMoveFolder}\\{fi.Name}", true);
                } catch (Exception ex) {
                    Log.Error(ex, $"変換元ファイル'{fi.Name}' を処理失敗フォルダ '{_settings.ErrorFileMoveFolder}' に移動できませんでした");
                }
                return false;
            }
            Log.Information($"ファイル'{fi.FullName}'処理終了");
        }

        return true;
    }

    /// <summary>
    /// 帳票の生成を行う
    /// </summary>
    private async Task<bool> MakeReport(FileInfo fi) {
        // 元ファイルの全行のデータ
        var reqData = new List<string[]>();

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

        //// ヘッダ行を取出し
        //if (_settings.HeaderRowCount > 0) {
        //    this._columnHeaders = reqData[_settings.HeaderRowCount - 1];
        //}

        // 読み込んだデータから帳票を作成する
        // 行ループ
        for (var rowNo = 0; rowNo < reqData.Count; rowNo++) {

            // 変換リスト
            var reportData = new List<ReportData>();
            //List<ReportData> reportData = new();
            //List<ReportData> reportData= new List<ReportData>;

            // クラスタ変換
            foreach (var fromTo in _settings.FromTo ) {

                //列番号取得1～
                int colNo = fromTo.ColumnNo - 1;

                // カラムデータを取出し
                
                string cellValue = reqData[rowNo][colNo];

                // カラムデータがない場合はエラー
                if (cellValue is null) {
                    Log.Error($"ファイル'{fi.FullName}'内の'{rowNo+1}'行目に列番号'{colNo+1}'列のデータが見つかりません");
                    return false;
                }

                // 日付変換
                var culture = new CultureInfo("ja-JP");

                // 日付文字列を変換して登録する
                // FromDateFormat
                if (fromTo.FromDateFormat is not null) {
                    DateTime dt;
                    // 日付文字列を変換して登録する
                    try {
                        dt = DateTime.ParseExact(cellValue, fromTo.FromDateFormat, null);
                    } catch (Exception ex) {
                        Log.Error(ex,$"CSVファイルの{colNo+1}列目の日付文字列'{cellValue}'または 設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'の FromDateFormat の指定'{fromTo.FromDateFormat}'が不正です");
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


                if (!string.IsNullOrEmpty(cellValue)) {
                    // データがあれば追加する
                    reportData.Add(new ReportData(fromTo, cellValue));
                }
            }

            // 設定情報が１つも無かった場合、行読出し終了
            if (reportData.Count==0) break;

            // 生成する帳票名を生成する
            //_repTopName = GetRepTopName(fi, reportData, reqData[row], row - _settings.HeaderRowCount);
            var name = ConvertNameFromFile(_settings.RepTopName, fi);
            name = ConvertNameFromReport(name);
            name = ConvertNameFromData(name, reportData, rowNo+1);
            _repTopName = name;

            // ラベル情報設定
            var label = ConvertNameFromFile(_settings.LabelName??"", fi);
            label = ConvertNameFromReport(label);
            label = ConvertNameFromData(label, reportData, rowNo+1);


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
            sb.AddLabels();

            // クラスタ変換
            foreach (var r in reportData.Where(m => m.fromto.ClusterID is not null)) {
                // SxxCxxを作成
                sb.Addクラスター(r.fromto.SheetNo??0, r.fromto.ClusterID??0);
            }
            // 備考変換
            foreach (var r in reportData.Where(m => m.fromto.RemarksNo is not null)) {
                // 備考を作成
                sb.Add備考(r.fromto.RemarksNo??0);
            }
            // システムキー変換
            foreach (var r in reportData.Where(m => m.fromto.SystemKeyNo is not null)) {
                // システムキーを作成
                sb.AddSystemKey(r.fromto.SystemKeyNo??0);
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
            foreach (var r in reportData.Where(m => m.fromto.ClusterID is not null)) {
                // SxxCxx
                sb.AddValue(r.data);
            }
            // 備考変換
            foreach (var r in reportData.Where(m => m.fromto.RemarksNo is not null)) {
                // 備考を作成
                sb.AddValue(r.data);
            }
            // システムキー変換
            foreach (var r in reportData.Where(m => m.fromto.SystemKeyNo is not null)) {
                // システムキーを作成
                sb.AddValue(r.data);
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

    /// <summary>
    /// 変換元ファイル読み込み（CSV）
    /// </summary>
    private List<string[]> GetRequestDataFromCsv(FileInfo fi) {
        var requestData = new List<string[]>();

        // CSVの設定
        var csvConf = new CsvConfiguration(CultureInfo.InvariantCulture) {
            Delimiter = _settings.Delimiter.ToLower() switch { "," => ",", "tab" => "\t", _ => "," },
            HasHeaderRecord = _settings.HasHeaderRecord,
            Encoding = _settings.Encode
        };

        using (var stream = new StreamReader(fi.FullName))
        using (var reader = new CsvReader(stream, csvConf)) {
            // ヘッダを読み飛ばす
            if (csvConf.HasHeaderRecord) {
                reader.Read();
            }
            // 行データを読み込む
            while (reader.Read()) {
                var datas = new List<string>();
                // 列データを読み込む
                for (int i = 0; i < reader.ColumnCount; i++) {
                    datas.Add(reader.GetField(i)??"");
                }
                requestData.Add(datas.ToArray());
            }
        }
        return requestData;
    }
}
