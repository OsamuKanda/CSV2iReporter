using ConMasIReporterLib.Models;
//using ExcelLib;
using Microsoft.Extensions.Configuration;
using System.Xml.Linq;
using Serilog;
using System.Text;
using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;
using ConMasIReporterLib;
using CSV2iReporter.Models;

namespace CSV2iReporter;
/// <summary>
/// i-Reporter自動帳票連携ソフト
/// </summary>
/// <remarks>
/// コンストラクタ
/// </remarks>
public class ReportMakeFromCSV(IConfigurationRoot config) {
    //private readonly Convert _convert = new(config.GetSection("Convert")).Get<Convert>();
    private readonly Models.Convert? _convert = config.GetSection("Convert").Get<Models.Convert>();
    private readonly ComMasWebAPIService2 _svc = new(config.GetSection("ConMasWeb"));

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
    //private string? _repTopName;
    ///// 列名称ヘッダ配列
    ///// </summary>
    //private string[]? _columnHeaders = null;


    /// <summary>
    /// クラスター定義が存在するか確認する
    /// </summary>
    /// <param name="cluster"></param>
    /// <param name="SheetNo"></param>
    /// <param name="ClusterID"></param>
    /// <returns></returns>
    static private bool HasCluster(IEnumerable<XElement>? clusters, int? SheetNo, int? ClusterID) {
        // クラスター
        if (clusters is not null) {
            // 定義されたシートNoとクラスターIDが存在するか？
            foreach (var repCluster in clusters) {
                string? sheetNo = repCluster?.Element("sheetNo")?.Value;
                string? clusterId = repCluster?.Element("clusterId")?.Value;

                if (!int.TryParse(sheetNo, out var sheet))
                    continue;

                if (!int.TryParse(clusterId, out var id))
                    continue;

                if ((!SheetNo.Equals(sheet)) || (!ClusterID.Equals(id)))
                    continue;

                //定義されたシートNoとクラスターIDが存在した
                return true;
            }
        }
        return false;
    }

    /// <summary>
    /// 変換元データ読込
    /// </summary>
    /// <param name="fi">変換元ファイル</param>
    /// <param name="enc">文字エンコード</param>
    /// <param name="delimiter">デリミタ(nullで自動判別）</param>
    /// <param name="hasHeaderRecord">ヘッダ行の有無</param>
    /// <returns></returns>
    private static List<string[]> GetRequestDataFromCsv(FileInfo fi, Encoding enc, string? delimiter, Boolean hasHeaderRecord) {

        var csvDatas = new List<string[]>();

        // CSVの形式を指定
        CsvConfiguration csvConf = new(CultureInfo.InvariantCulture) {
            HasHeaderRecord = hasHeaderRecord,
            Encoding = enc,
            Delimiter = delimiter ?? ","
        };

        // CSVを読み出す
        using (var stream = new StreamReader(fi.FullName, enc)) {
            using var reader = new CsvReader(stream, csvConf);
            string detectedDelimiter = reader.Configuration.Delimiter switch { "\t" => "tab", _ => reader.Configuration.Delimiter };
            //Log.Information($"ファイル'{fi.FullName}'内のデリミタは'{detectedDelimiter}'で認識されました");
            // ヘッダを読み飛ばす
            //// ヘッダがある場合は先頭にヘッダ行を追加する
            //if (csvConf.HasHeaderRecord) {
            //    reader.Read();
            //    reader.ReadHeader();
            //    var datas = new List<string>();
            //    for (int i = 0; i < reader.HeaderRecord?.Length; i++) {
            //        datas.Add(reader.HeaderRecord[i] ?? "");
            //    }
            //    csvDatas.Add([.. datas]);
            //}

            // 行データを読み込む
            while (reader.Read()) {
                var datas = new List<string>();
                // 列データを読み込む
                for (int i = 0; i < reader.ColumnCount; i++) {
                    datas.Add(reader.GetField(i) ?? "");
                }
                csvDatas.Add([.. datas]);
            }
        }
        return csvDatas;
    }

    /// <summary>
    /// 初期化
    /// </summary>
    public async Task<bool> Init() {

        if (!File.Exists($@"{AutoReportLib.Init.SettingFileFullPath}")) {
            Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'が ありません");
            return false;
        }

        // 設定ファイルにConvertセクションが無い
        if (_convert is null) {
            Log.Error($"設定ファイルに {nameof(Models.Convert)} セクションの記述がありません");
            return false;

        }
        // 設定ファイルの帳票定義チェック
        if ( (_convert.DefTopId is null) && (_convert.DefTopName is null) ) {
            Log.Error($"設定ファイルの {nameof(Models.Convert)} セクションに {nameof(_convert.DefTopId)} または {nameof(_convert.DefTopName)} の記述がありません");
            return false;
        }

        _defTopName = _convert.DefTopName;
        _defTopId = _convert.DefTopId;

        if(_defTopId is null) {

            // 名称指定の場合、最新の帳票IDを取得する

            // 指定の定義情報を読み込む
            var reports = await _svc.Req定義一覧取得Async(_convert.DefTopName??"");

            if ((reports is not null) && (!reports.IsEmpty)) {
                // 帳票定義名が一致する帳票定義から最新の更新日時の帳票定義IDを取得する
                // （更新日の降順に帳票定義を並べ替えて最初の帳票IDを取得）
                _defTopId = (from item in reports.Elements("items").Elements("item")
                             where (string?)item.Element("name") == _defTopName
                             let updateDateTime = (DateTime?)item.Element("updateTime") 
                             let id = (int?)item.Element("itemId") 
                             orderby updateDateTime descending
                             select id).First();
             }
            if (_defTopId is null) {
                Log.Error($"帳票定義名称が'{_convert.DefTopName}'の帳票定義が見つかりません");
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

        //// 帳票備考のチェック
        if (_convert.Remarks is not null) {
            foreach (var remark in _convert.Remarks) {
                if((remark.No < 1) || (remark.No > 10)) {
                    Log.Error($"帳票備考のIndexに'{remark.No}'が指定されました。帳票備考のIndexには1～10迄を指定してください");
                    return false;
                }
            }
        }
        //// システムキーのチェック
        if (_convert.SystemKeys is not null) {
            foreach (var systemKey in _convert.SystemKeys) {
                if ((systemKey.No < 1) || (systemKey.No > 5)) {
                    Log.Error($"システムキーのIndexに'{systemKey.No}'が指定されました。システムキーのIndexには1～5迄を指定してください");
                    return false;
                }
            }
        }
        //// クラスターのチェック
        if (_convert.Sheet is not null) {
            foreach (var sheet in _convert.Sheet) {
                foreach(var cluster in sheet.Cluster) {
                    // クラスター
                    var clusters = conmas.Element("detailInfo")?.Element("clusters")?.Elements("cluster");
                    if (!HasCluster(clusters, sheet.No, cluster.Id)) {
                        Log.Error($"設定ファイルの {nameof(sheet)}'{sheet.No}'、{nameof(cluster)}'{cluster.Id}' は 帳票定義'{_defTopId}' にありません");
                        return false;
                    }
                }
            }
        }

        // 監視、正常時、異常時データ格納先のディレクトリが無ければ生成
        if( !Directory.Exists(_convert.InputPath)) {
            try {
                Directory.CreateDirectory(_convert.InputPath);
            } catch (Exception ex) {
                Log.Error(ex,$"設定ファイルの {nameof(_convert.InputPath)}'{_convert.InputPath}'を作成できませんでした");
            }
        }
        if (_convert.OutputPath is not null) {
            if (!Directory.Exists(_convert.OutputPath)) {
                try {
                    Directory.CreateDirectory(_convert.OutputPath);
                } catch (Exception ex) {
                    Log.Error(ex, $"設定ファイルの {nameof(_convert.OutputPath)}'{nameof(_convert.OutputPath)}'を作成できませんでした");
                }
            }
        }
        if (!Directory.Exists(_convert.SuccessFileMovePath)) {
            try {
                Directory.CreateDirectory(_convert.SuccessFileMovePath);
            } catch (Exception ex) {
                Log.Error(ex,$"設定ファイルの {nameof(_convert.SuccessFileMovePath)}'{_convert.SuccessFileMovePath}'を作成できませんでした");
            }
        }
        if (!Directory.Exists(_convert.ErrorFileMovePath)) {
            try {
                Directory.CreateDirectory(_convert.ErrorFileMovePath);
            } catch (Exception ex) {
                Log.Error(ex,$"設定ファイルの {nameof(_convert.ErrorFileMovePath)}'{_convert.ErrorFileMovePath}'を作成できませんでした");
            }
        }

        return true;
    }

    /// <summary>
    /// 文字列に指定のキー文字列があるか確認する　{key:format}
    /// </summary>
    /// <param name="value"></param>
    /// <param name="key"></param>
    /// <returns>   キー文字列が無かった場合nullを返す
    ///             キー文字列は存在したが、フォーマット指定が無い場合、空白文字列を返す
    ///             キー文字列と対応するフォーマットがあった場合はフォーマット文字列を返す</returns>
    static private string? GetKeyFormat(string value, string key) {

        // "{Key}"だけの設定があれば
        if (value.Contains($"{{{key}}}", StringComparison.CurrentCulture)) {
            return "";
        }

        // "{Key"が無ければリターン
        var sp = value.IndexOf($"{{{key}:", StringComparison.CurrentCulture);
        if (sp == -1) {
            return null;
        }

        // キー文字列の長さを取得
        var keyLength = ($"{{{key}:").Length;

        // "}"が無ければリターン
        var len = value[(sp+keyLength)..].IndexOf('}', StringComparison.CurrentCulture);
        if (len == -1) {
            return null;
        }

        // フォーマット文字列を取り出す
        return value.Substring(sp+keyLength, len);
    }

    /// <summary>
    /// 数値フォーマット書式付きキーを変換する
    /// </summary>
    /// <param name="value">変換元文字列</param>
    /// <param name="key">検索キー</param>
    /// <param name="dt">変換したい日時データ</param>
    /// <returns></returns>
    static private string FormatDecimal(string value, string key, Decimal dt) {
        var format = GetKeyFormat(value, key);

        //キーが無ければリターンする
        if (format is null) {
            return value;
        }

        // フォーマットが指定されていなければデフォルトの書式を採用する
        if (format == "") {
            return value.Replace($"{{{key}}}", dt.ToString("d"), StringComparison.CurrentCultureIgnoreCase);
        } else {
            return value.Replace($"{{{key}:{format}}}", dt.ToString(format), StringComparison.CurrentCultureIgnoreCase);
        }
    }

    /// <summary>
    /// 日付フォーマット書式付きキーを変換する
    /// </summary>
    /// <param name="value">変換元文字列</param>
    /// <param name="key">検索キー</param>
    /// <param name="dt">変換したい日時データ</param>
    /// <returns></returns>
    private static string FormatDateTime(string value, string key, DateTime dt) {
        var format = GetKeyFormat(value, key);

        //キーが無ければリターンする
        if (format is null) {
            return value;
        }

        // フォーマットが指定されていなければデフォルトの書式を採用する
        if (format == "") {
            return value.Replace($"{{{key}}}", dt.ToString("yyyy/MM/dd HH:mm:ss"), StringComparison.CurrentCultureIgnoreCase);
        } else {
            return value.Replace($"{{{key}:{format}}}", dt.ToString(format), StringComparison.CurrentCultureIgnoreCase);
        }
    }
    /// <summary>
    /// ファイル情報を元に文字列変換を行う
    /// </summary>
    /// <param name="value"></param>
    /// <param name="fi"></param>
    /// <returns></returns>
    private static string ConvertNameFromFile(string value, FileInfo fi) {

        // 元ファイル名（ベース）
        value = value.Replace("{FileName}", Path.GetFileNameWithoutExtension(fi.Name), StringComparison.CurrentCultureIgnoreCase);

        // 依頼ファイルのタイムスタンプ（更新日時）
        value = FormatDateTime(value, "FileDate", fi.LastWriteTime);
        return value;
    }
    /// <summary>
    /// 情報を元に文字列変換を行う
    /// </summary>
    /// <param name="value"></param>
    /// <param name="fi"></param>
    /// <returns></returns>
    private static string ConvertNameFromNow(string value) {

        // 現在のタイムスタンプ（更新日時）
        return FormatDateTime(value, "Now", DateTime.Now);

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
        repTopName = repTopName.Replace("{defTopName}", _defTopName??"", StringComparison.CurrentCultureIgnoreCase);

        return repTopName;
    }

    /// <summary>
    /// 変換ルールに従って文字列を変換する
    /// </summary>
    /// <param name="value">変換値</param>
    /// <param name="fi">変換元CSVファイルの情報</param>
    /// <param name="rowNo">CSVファイル内の処理行番号</param>
    /// <param name="rowData">CSVファイル内の１行のデータ</param>
    /// <param name="maxColumnCount">CSVファイルの最大列数</param>
    /// <param name="headers">CSVヘッダ文字列</param>
    /// <returns></returns>
    private string GetFormatData(string value, FileInfo fi,int rowNo, string[] rowData, int maxColumnCount, string[]? headers) {

        // 帳票ID、ユーザー名による変換
        value = ConvertNameFromReport(value);

        // ファイル情報による変換
        value = ConvertNameFromFile(value, fi);

        // 処理現在日時による変換
        value = ConvertNameFromNow(value);

        //// 処理行番号による変換
        value = FormatDecimal(value, "RowNo", rowNo);

        // クラスタ定義が存在するか？（起動時ににチェックしているのであるはず）
        if (_convert is null) {
            return value;
        }

        /// ローカル関数
        string  ReplaceColumn(string value, string key, string replace, int colNo) {
            // 指定のデータ列の指定が存在するか？
            var format = GetKeyFormat(value, key);

            // 無ければ次の列番号へ
            if (format is null) {
                return value;
            } else if (format == "") {
                // 出力フォーマット指定が無い場合は、値をそのまま登録する
                value = value.Replace("{" + key + "}", replace);
            } else {
                // 出力フォーマットの指定がある場合

                // Csv桁変換データがあれば桁変換を探す
                if (_convert is null) {
                    return value;
                }
                if (_convert.CsvColumn is null) {
                    return value;
                }
                DateTime? dte = null;
                Decimal? num = null;
                foreach (var csvColumn in _convert.CsvColumn.Where(m => m.No == (colNo + 1))) {
                    if (csvColumn.DateFormat is not null) {
                        // 変換設定を元に入力日付を取り出す
                        dte = DateTime.ParseExact(replace, csvColumn.DateFormat, null);
                        // 出力フォーマットの指定がある場合
                        value = value.Replace("{" + key + ":" + format + "}", ((DateTime)dte).ToString(format));
                        break;
                    } else {
                        // 変換設定を元に数値として取り出す
                        num = Decimal.Parse(replace);
                        // 変換元フォーマットが指定されている場合は、変換されたDecimal値を書式変換する
                        value = value.Replace("{" + key + ":" + format + "}", ((Decimal)num).ToString(format));
                    }
                    break;
                }
            }
            return value;
        }
        // 列番号指定データの変換
        int colNo = 0;

        for (; colNo < rowData.Length; colNo++) {
            value = ReplaceColumn(value, (colNo + 1).ToString(), rowData[colNo], colNo);
        }
        // CSVデータに存在しない列は空白に置き換える
        for (; colNo < maxColumnCount; colNo++) {
            // 指定のデータ列の指定が存在するか？
            value = ReplaceColumn(value, (colNo + 1).ToString(), "", colNo);
        }

        // ヘッダー文字列で変換する
        if( headers is not null) {
            // ヘッダー指定データの変換
            colNo = 0;

            for (; colNo < rowData.Length; colNo++) {
                value = ReplaceColumn(value, headers[colNo], rowData[colNo], colNo);
            }
            // CSVデータに存在しない列は空白に置き換える
            for (; colNo < maxColumnCount; colNo++) {
                value = ReplaceColumn(value, headers[colNo], "", colNo);
            }
        }

        return value;
    }

    /// <summary>
    /// 帳票生成実行
    /// </summary>
    public async Task<bool> Execute() {

        // 設定ファイルにConvertセクションが無い
        if (_convert is null) {
            Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'に {nameof(Models.Convert)} セクションの記述がありません");
            return false;
        }

        // 監視ディレクトリ検索
        var files = Directory.GetFiles(_convert.InputPath, _convert.InputFilePattern).OrderBy(f => File.GetLastWriteTime(f));

        // 見つかったファイル分の生成処理を行う
        foreach ( var file in files ) {
            var fi = new FileInfo(file);

            Log.Information($"ファイル'{fi.FullName}'処理開始");

            // 帳票の生成
             var ret = await MakeReport(fi);

            // 成功、失敗にによるファイルの移動
            if (ret != "") {

                // XMLファイル出力の指定があれば出力する
                if (_convert.OutputPath is not null) {
                    var outputFileName = ConvertNameFromFile(_convert.OutputFileName ?? "{DefTopName}_{FileDate:yyyyMMddHHmmss}.csv", fi);
                    outputFileName = ConvertNameFromReport(outputFileName);
                    outputFileName = ConvertNameFromNow(outputFileName);
                    var targetFile = Path.Combine(_convert.OutputPath, outputFileName);
                    if (File.Exists(targetFile)) {
                        Log.Information($"XML Outputファイル['{outputFileName}']が既に存在するため上書きします");
                    }
                    var utf8Encoding = System.Text.Encoding.UTF8;
                    using var streamWriter = new StreamWriter(targetFile, false, utf8Encoding);
                    await streamWriter.WriteAsync(ret);
                }

                Log.Information($"正常処理ファイル'{fi.FullName}'移動先'{_convert.SuccessFileMovePath}\\{fi.Name}'");
                try {
                    File.Move(fi.FullName, $"{_convert.SuccessFileMovePath}\\{fi.Name}", true);
                } catch (Exception ex) {
                    Log.Error(ex, $"変換元ファイル'{fi.Name}' を処理正常フォルダ '{_convert.SuccessFileMovePath}' に移動できませんでした");
                    return false;
                }
            } else {
                Log.Information($"異常処理ファイル'{fi.FullName}'移動先'{_convert.ErrorFileMovePath}\\{fi.Name}'");
                try {
                    File.Move(fi.FullName, $"{_convert.ErrorFileMovePath}\\{fi.Name}", true);
                } catch (Exception ex) {
                    Log.Error(ex, $"変換元ファイル'{fi.Name}' を処理失敗フォルダ '{_convert.ErrorFileMovePath}' に移動できませんでした");
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
    private async Task<string> MakeReport(FileInfo fi) {

        // 設定ファイルにConvertセクションが無い
        if (_convert is null) {
            Log.Error($"設定ファイル'{AutoReportLib.Init.SettingFileFullPath}'に {nameof(Models.Convert)} セクションの記述がありません");
            return "";
        }

        // CSVファイルの文字エンコードを設定する
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        //エンコードを設定
        Encoding enc = ((_convert.Encode ?? "utf-8").ToLower() switch {
            "sjis" or "shift-jis" or "shift_jis" => Encoding.GetEncoding("Shift-JIS"),
            "utf8" or "utf-8" or "utf_8" => Encoding.UTF8,
            "unicode" or "utf-16" or "utf_16" => Encoding.Unicode,
            "euc-jp" or "euc_jp" or "utf_16" => Encoding.GetEncoding("euc-jp"),
            "iso-2022-jp" or "iso_2022_jp" or "jis" => Encoding.GetEncoding("iso-2022-jp"),
            _ => Encoding.UTF8
        });

        // CSVファイルのデリミタを設定（自動検出はnull）
        string? delimiter = _convert.Delimiter.ToLower() switch { "tab" 
            => "\t", "auto" => null, _ => _convert.Delimiter.ToLower() };

        // 元ファイルの全行のデータ
        var allCsvData = new List<string[]>();

        // CSVファイルからデータを読み込む
        allCsvData = GetRequestDataFromCsv( fi, enc, delimiter, _convert.HasHeaderRecord );

        // 最大列数を取得する
        int maxColumnCount = 0;
        foreach(var row in allCsvData){
            if (row.Length > maxColumnCount) {
                maxColumnCount = row.Length;
            }
        }

        // ヘッダ行取得
        string[]? headers = null;

        // CSVヘッダ処理
        if (_convert.HasHeaderRecord) {
            // CSVからヘッダ行を取得する
            headers = allCsvData[0];
            // CSVヘッダ行を削除する
            allCsvData.RemoveAt(0);
        }

        // 自動帳票作成クラスを作成
        Models.AutoReport conmas = new();

        // 読み込んだデータから帳票を作成する
        for (var rowNo = 0; rowNo < allCsvData.Count; rowNo++) {

            // 帳票名設定
            Models.AutoReport.Top top = new(_defTopId ?? 0) {
                repTopName = GetFormatData( _convert.RepTopName??"{DefTopName}_{RowNo}_{Now:yyyyMMddHHmmss}", fi, rowNo, allCsvData[rowNo], maxColumnCount, headers)
            };

            // 帳票備考変換
            if (_convert.Remarks is not null) {
                foreach (var def in _convert.Remarks) {
                    top.remarks[def.No - 1] = GetFormatData(def.Output, fi, rowNo, allCsvData[rowNo], maxColumnCount, headers);
                }
            }
            // システムキー変換
            if (_convert.SystemKeys is not null) {
                foreach (var def in _convert.SystemKeys) {
                    top.systemKey[def.No - 1] = GetFormatData(def.Output, fi, rowNo, allCsvData[rowNo], maxColumnCount, headers);
                }
            }

            // ラベル情報変換
            string? label = null;
            if (_convert.LabelName is not null) {
                label = GetFormatData(_convert.LabelName, fi, rowNo, allCsvData[rowNo], maxColumnCount, headers);
                // ラベル情報設定（セミコロンで分解して設定）
                if (label is not null) {
                    foreach (var s in label.Split(";")) {
                        top.label.Add(new Models.AutoReport.Label(s));
                    }
                }
            }

            // クラスタ変換
            if (_convert.Sheet is not null) {
                foreach (var s in _convert.Sheet) {
                    // シート設定オブジェクト作成
                    Models.AutoReport.Top.Sheet dst = new(s.No);
                    foreach (var c in s.Cluster) {
                        //クラスタデータ変換
                        var output = GetFormatData(c.Output, fi, rowNo, allCsvData[rowNo], maxColumnCount, headers);
                        Models.AutoReport.Top.Sheet.Cluster cluster = new(s.No, c.Id, output);
                        dst.cluster.Add(cluster);
                    }
                    top.sheet.Add(dst);
                }
            }
    
            // 自動帳票作成依頼データに追加
            conmas.top.Add(top);
        }

        // 自動帳票データをシリアライズしてXML文字列に変換
        System.Xml.Serialization.XmlSerializer xml = new(conmas.GetType());
        StringWriter mem = new();
        xml.Serialize(mem, conmas);
        XDocument doc = XDocument.Parse(mem.ToString(),LoadOptions.PreserveWhitespace);
        var xmlString = doc.Element("AutoReport")?.Element("conmas")?.ToString(SaveOptions.None);

        // iReporterサーバへ自動帳票作成要求を行う
        var reqresult = await _svc.Req自動帳票作成Xml2Async(xmlString??"");

        // 要求が失敗した場合は中止する
        if (!reqresult.IsSuccess) {
            Log.Error($"自動帳票作成要求に失敗しました。エラーコードは({reqresult.ErrorCode})です");
            return "";
        }

         return xmlString??"";
    }

}
