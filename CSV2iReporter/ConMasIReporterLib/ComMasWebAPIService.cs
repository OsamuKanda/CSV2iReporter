//#define ComMasWebAPIへの要求はWebClientを使用する
using ConMasIReporterLib.Models;
using Serilog;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace ConMasIReporterLib;

public class ComMasWebAPIService {
    public bool 帳票の状態を入力完了にする { get; set; } = false;

    /// <summary>ConMas Server アドレス</summary>
    protected readonly string _url;
    /// <summary>ConMas Server ユーザーID</summary>
    protected readonly string _user;
    /// <summary>ConMas Server パスワード</summary>
    protected readonly string _password;

    /// <summary>プロキシ サーバー</summary>
    protected readonly string? _proxyServer;
    /// <summary>プロキシ ユーザーID</summary>
    protected readonly string? _proxyUserName;
    /// <summary>プロキシ パスワード</summary>
    protected readonly string? _proxyPassword;
    /// <summary>プロキシ ドメイン</summary>
    protected readonly string? _proxyDomain;

    protected readonly IConMasWebAPIClient _client;

    /// <summary>
    /// コンストラクタ
    /// </summary>
    /// <param name="url"></param>
    /// <param name="user"></param>
    /// <param name="password"></param>
    public ComMasWebAPIService(string url, string user, string password, string? proxyServer = null, string? proxyUserName = null, string? proxyPassword = null, string? proxyDomain = null) {
        _url = url;
        _user = user;
        _password = password;
        _proxyServer = proxyServer;
        _proxyUserName = proxyUserName;
        _proxyPassword = proxyPassword;
        _proxyDomain = proxyDomain;

        _client =
#if ComMasWebAPIへの要求はWebClientを使用する
        new ConMasWebAPIClient();   
#else
        new ConMasHttpClient(url, user, password, proxyServer, proxyUserName, proxyPassword, proxyDomain);
#endif
    }

    public async Task<bool> LoginAsync() {
        // ログインする
        var result = await _client.LoginAsync(_url, _user, _password);
        Log.Information($"ConMasWebAPI Login {(result ? "成功" : "失敗")}");
        return result;
    }
    public async Task LogoutAsync() {
        // ログアウトする
        var resultLogout = await _client.LogoutAsync(_url);
        Log.Information($"ConMasWebAPI Logout {(resultLogout ? "成功" : "失敗")}");
    }

    /// <summary>
    /// 帳票更新要求を行います
    /// </summary>
    /// <param name="data">簡易CSV形式のデータ(本メソッド内でToString()呼び出しするため、
    /// StringBuilderなどを直接渡してもよいこととします)</param>
    /// <returns></returns>
    public async Task 帳票更新Async(object data) {
        if (!await LoginAsync()) {
            return;
        }

        // 帳票更新要求を行う
        var resultUpdateReport = new ConMasWebAPIResult(await _client.UpdateReportAsync(
            _url,
            "csv",
            Encoding.UTF8,
            false,
            data?.ToString() ?? "",
            "upload.csv"
            ));

        if (resultUpdateReport.IsSuccess) {
            Log.Information($"ConMasWebAPI UpdateReport 成功");

            // 帳票更新要求の応答から更新された帳票IDを取得する
            var results = resultUpdateReport.ReportResults();
            if (results is not null) {
                var logstr = new StringBuilder();
                foreach (var result in results) {
                    var topId = result.TopId;
                    if (topId is not null) {
                        logstr.Append($" {topId}");
                    }
                }
                Log.Information($"　→更新帳票ID = {logstr}");
            }
        } else {
            Log.Information($"ConMasWebAPI UpdateReport 失敗({resultUpdateReport.ErrorCode})");
        }

        await LogoutAsync();
    }

    /// <summary>
    /// 編集ステータスを入力完了とするXMLファイル作成する
    /// </summary>
    /// <param name="list作成ID">対象repTopIdのリスト</param>
    public static FileInfo 編集ステータスを入力完了とするXMLファイル作成(IList<int> list作成ID) {
        // 帳票更新のXMLファイルを作成する
        var xmlFi = Create帳票更新XMLファイル情報();
        Log.Information($"帳票更新XMLファイル = {xmlFi.Name}");

        using var xw = XmlWriter.Create(xmlFi.FullName);

        //ルートノード作成
        xw.WriteStartElement("conmas");

        //
        foreach (var repTopId in list作成ID) {
            xw.WriteStartElement("top");
            xw.WriteElementString("repTopId", $"{repTopId}");
            xw.WriteElementString("editStatus", "4");//4:入力完了
            xw.WriteEndElement();
        }

        return xmlFi;
    }

    /// <summary>
    /// 簡易CSVファイルのファイル名を生成する
    /// アプリケーションのディレクトリに「REQCSV」フォルダを作成し、同フォルダ内にXMLファイルを作成します。
    /// </summary>
    protected static FileInfo Create帳票更新XMLファイル情報() {
        var basepath = AppDomain.CurrentDomain.BaseDirectory;

        var dirnam = $"{basepath}\\REQCSV";
        Directory.CreateDirectory(dirnam);
        var filnam = $"{dirnam}\\{DateTime.Now:yyyyMMddHHmmssfff}.xml";

        return new FileInfo(filnam);
    }

    /// <summary>
    /// 自動帳票の作成を行う(簡易CSV_ZIP)
    /// </summary>
    /// <param name="fnam"></param>
    public async Task Req自動帳票作成Async(string fnam, string type) {
        if (!await LoginAsync()) {
            return;
        }

        // 自動帳票作成要求を行う
        var resultAutoGenerate = new ConMasWebAPIResult(await _client.AutoGenerateAsync(_url
            , type
            , System.Text.Encoding.UTF8
            , 1//1"を指定する事で、アップロードファイル中の"作成ユーザー"のユーザーが帳票登録者として登録される。
            , 1//"1"を指定する事で、アップロードファイル中の"ラベル"が階層設定されていた場合に最下層のラベルのみに紐づけられます。
            , fnam
            ));

        // 自動帳票作成要求の応答から生成された帳票IDを取得する
        var list作成ID = GetList作成ID(resultAutoGenerate);
        if (帳票の状態を入力完了にする) { // 帳票の状態を入力完了にしない
            if (list作成ID.Count != 0) {
                await 作成した自動帳票を入力完了にするAsync(list作成ID);
            }
        }
        await LogoutAsync();
    }

    public async Task Req自動帳票作成CsvZipAsync(string fnam) =>
        await Req自動帳票作成Async(fnam, "csvZipSimple");

    public async Task Req自動帳票作成CsvAsync(string fnam) =>
        await Req自動帳票作成Async(fnam, "csvSimple");

    /// <summary>
    /// ファイルを作成せずに送信
    /// </summary>
    /// <param name="fnam"></param>
    /// <returns></returns>
    public async Task<ConMasWebAPIResult> Req自動帳票作成Csv2Async(string data) {
        if (!await LoginAsync()) {
            return new ConMasWebAPIResult(null);
        }

        // 自動帳票作成要求を行う
        var resultAutoGenerate = new ConMasWebAPIResult(await _client.AutoGenerateAsync(_url
            , "csvSimple"
            , Encoding.UTF8
            , 1//1"を指定する事で、アップロードファイル中の"作成ユーザー"のユーザーが帳票登録者として登録される。
            , 1//"1"を指定する事で、アップロードファイル中の"ラベル"が階層設定されていた場合に最下層のラベルのみに紐づけられます。
            , data
            , "upload.csv"
            ));

        // 自動帳票作成要求の応答から生成された帳票IDを取得する
        var list作成ID = GetList作成ID(resultAutoGenerate);
#if false // 帳票の状態を入力完了にしない
        if (list作成ID.Any()) {
            await 作成した自動帳票を入力完了にするAsync(list作成ID);
        }
#endif
        await LogoutAsync();

        return resultAutoGenerate;
    }

    /// <summary>
    /// ファイルを作成せずに送信
    /// </summary>
    /// <param name="fnam"></param>
    /// <returns></returns>
    public async Task<ConMasWebAPIResult> Req自動帳票作成Xml2Async(string data) {
        if (!await LoginAsync()) {
            return new ConMasWebAPIResult(null);
        }

        // 自動帳票作成要求を行う
        var resultAutoGenerate = new ConMasWebAPIResult(await _client.AutoGenerateAsync(_url
            , "xml"
            , Encoding.UTF8
            , 1//1"を指定する事で、アップロードファイル中の"作成ユーザー"のユーザーが帳票登録者として登録される。
            , 1//"1"を指定する事で、アップロードファイル中の"ラベル"が階層設定されていた場合に最下層のラベルのみに紐づけられます。
            , data
            , "upload.csv"
            ));

        // 自動帳票作成要求の応答から生成された帳票IDを取得する
        var list作成ID = GetList作成ID(resultAutoGenerate);
#if false // 帳票の状態を入力完了にしない
        if (list作成ID.Any()) {
            await 作成した自動帳票を入力完了にするAsync(list作成ID);
        }
#endif
        await LogoutAsync();

        return resultAutoGenerate;
    }

    /// <summary>
    /// 帳票ファイル取得
    /// </summary>
    /// <param name="帳票ID"></param>
    /// <param name="fileType"></param>
    /// <param name="fileName"></param>
    /// <returns></returns>
    public async Task<(bool result, byte[] fileData)> Req帳票ファイル取得Async(int 帳票ID, string fileType, string fileName) {
        ArgumentNullException.ThrowIfNull(fileName);

        if (!await LoginAsync()) {
            return (false, Array.Empty<byte>());
        }

        // 自動帳票作成要求を行う
        var resultGetReportFile = await _client.GetReportFileAsync(
            _url
            , fileType
            , 帳票ID
            );

        await LogoutAsync();

        return resultGetReportFile;
    }

    /// <summary>
    /// 自動帳票作成要求の応答から生成された帳票IDを取得する
    /// </summary>
    /// <param name="resultAutoGenerate">自動帳票作成の応答</param>
    /// <returns></returns>
    private static List<int> GetList作成ID(ConMasWebAPIResult resultAutoGenerate) {
        var list作成ID = new List<int>();
        if (resultAutoGenerate.IsSuccess) {
            Log.Information($"ConMasWebAPI AutoGenerate 成功");

            var results = resultAutoGenerate.ReportResults();
            if (results is not null) {
                var logstr = new StringBuilder();
                foreach (var result in results) {
                    var topId = result.TopId;
                    if (topId is not null) {
                        list作成ID.Add(int.Parse(topId));
                        logstr.Append($" {topId}");
                    }
                }
                Log.Information($"　→自動作成帳票ID = {logstr}");
            }
        } else {
            Log.Information($"ConMasWebAPI AutoGenerate 失敗({resultAutoGenerate.ErrorCode})");
        }
        return list作成ID;
    }

    // 作成した自動帳票を「入力完了」にする
    private async Task 作成した自動帳票を入力完了にするAsync(IList<int> list作成ID) {
        // 帳票更新のXMLファイルを作成する
        var xmlFi = 編集ステータスを入力完了とするXMLファイル作成(list作成ID);

        // 帳票更新要求を行う
        var resultUpdateReport = new ConMasWebAPIResult(await _client.UpdateReportAsync(
            _url
            , "xml"
            , "65001"
            , true
            , xmlFi.FullName
            ));

        if (resultUpdateReport.IsSuccess) {
            Log.Information($"ConMasWebAPI UpdateReport 成功");

            // 帳票更新要求の応答から更新された帳票IDを取得する
            var results = resultUpdateReport.ReportResults();
            if (results is not null) {
                var logstr = new StringBuilder();
                foreach (var result in results) {
                    if (result.TopId is not null) {
                        logstr.Append($" {result.TopId}");
                    }
                }
                Log.Information($"　→更新帳票ID = {logstr}");
            }
        } else {
            Log.Information($"ConMasWebAPI UpdateReport 失敗({resultUpdateReport.ErrorCode})");
        }
    }

    /// <summary>
    /// 定義一覧取得
    /// </summary>
    /// <returns></returns>
    public async Task<XElement?> Req定義一覧取得Async(string reportName = "") {
        if (!await LoginAsync()) {
            return null;
        }

        // 自動帳票作成要求を行う
        var resultGetReportFile = await _client.GetDefinitionListAsync(_url, reportName);

        await LogoutAsync();

        return resultGetReportFile;
    }

    /// <summary>
    /// 定義簡易詳細情報取得
    /// </summary>
    /// <param name="topId"></param>
    /// <returns></returns>
    public async Task<XElement?> Req定義簡易詳細情報取得Async(int topId) {
        if (!await LoginAsync()) {
            return null;
        }

        // 自動帳票作成要求を行う
        var resultGetReportFile = await _client.GetDefinitionDetailAsync(
            _url,
            $"{topId}"
            );

        await LogoutAsync();

        return resultGetReportFile;
    }
}
