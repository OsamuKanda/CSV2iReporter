using System.Collections.Specialized;
using System.Net;
using System.Text;
using System.Xml.Linq;

namespace ConMasIReporterLib;

/// <summary>
/// 拡張WebClientクラス（Cookie対応）
/// <para>
/// Cookieによるセッション管理に対応したWebClienクラス
/// </para>
/// </summary>
#pragma warning disable SYSLIB0014 // 型またはメンバーが旧型式です
public class WebClientEx : System.Net.WebClient {
    private readonly CookieContainer _cookieContainer = new();

    protected override WebRequest GetWebRequest(Uri url) {
        var webRequest = base.GetWebRequest(url);
        if (webRequest is HttpWebRequest request) {
            // コンテナーを設定することで、Cookie の取得や送信をやってもらう。
            request.CookieContainer = _cookieContainer;
        }
        return webRequest;
    }

}
#pragma warning restore SYSLIB0014 // 型またはメンバーが旧型式です

public class ConMasWebClient : IDisposable, IConMasWebAPIClient {
    /// <summary>
    /// Web Client
    /// </summary>
    private readonly WebClientEx webClient;

    /// <summary>
    /// コンストラクタ
    /// </summary>
    public ConMasWebClient() {
        webClient = new WebClientEx();

    }

    /// <summary>
    /// デストラクタ
    /// </summary>
    public void Dispose() {
        webClient.Dispose();
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// ログイン
    /// </summary>
    /// <param name="url"></param>
    /// <param name="userId"></param>
    /// <param name="password"></param>
    /// <returns></returns>
    public async Task<bool> LoginAsync(string url, string userId, string password) {
        await Task.Delay(0);
        var param = new NameValueCollection {
            { "command", "Login" },
            { "user", userId },
            { "password", password }
        };

        var result = webClient.UploadValues(url, param);
        var resultXml = XDocument.Parse(Encoding.UTF8.GetString(result));
        var code = resultXml?.Element("conmas")?.Element("loginResult")?.Element("code")?.Value;
        return code == "0";
    }

    /// <summary>
    /// ログアウト
    /// </summary>
    /// <param name="url"></param>
    /// <returns></returns>
    public async Task<bool> LogoutAsync(string url) {
        await Task.Delay(0);
        var param = new NameValueCollection {
            { "command", "Logout" }
        };

        var result = webClient.UploadValues(url, param);
        var resultXml = XDocument.Parse(Encoding.UTF8.GetString(result));
        var code = resultXml?.Element("conmas")?.Element("loginResult")?.Element("code")?.Value;
        return code == "0";
    }

    /// <summary>
    /// 自動帳票作成
    /// </summary>
    /// <param name="url"></param>
    /// <param name="type"></param>
    /// <param name="encoding"></param>
    /// <param name="userMode"></param>
    /// <param name="labelMode"></param>
    /// <param name="uploadFilePath"></param>
    /// <returns></returns>
    public async Task<XDocument> AutoGenerateAsync(string url, string type, Encoding encoding, int userMode, int labelMode, string uploadFilePath, bool calculateEnable = true) {
        await Task.Delay(0);
        var strEncoding = encoding == Encoding.UTF8 ? "65001" : "932";
        var urlString = $"{url}?command=AutoGenerate&type={type}&encoding={strEncoding}&userMode={userMode}&labelMode={labelMode}{(calculateEnable ? "&calculateEnable=1" : "")}";
        var result = webClient.UploadFile(urlString, uploadFilePath);
        return XDocument.Parse(Encoding.UTF8.GetString(result));
    }

    /// <summary>
    /// 定義一覧取得
    /// </summary>
    /// <param name="url"></param>
    /// <returns></returns>
    public XElement GetDefinitionList(string url) {
        var param = new NameValueCollection {
            { "command", "GetDefinitionList" }
        };
        var result = webClient.UploadValues(url, param);
        return XElement.Parse(Encoding.UTF8.GetString(result));
    }

    /// <summary>
    /// 帳票更新
    /// </summary>
    /// <param name="url"></param>
    /// <param name="type"></param>
    /// <param name="encoding"></param>
    /// <param name="uploadFilePath"></param>
    /// <returns></returns>
    public async Task<XDocument> UpdateReportAsync(string url, string type, string encoding, bool enableUpdateUser, string uploadFilePath) {
        await Task.Delay(0);
        var urlString = $"{url}?command=UpdateReport&type={type}&encoding={encoding}";
        var result = webClient.UploadFile(urlString, uploadFilePath);
        return XDocument.Parse(Encoding.UTF8.GetString(result));
    }

    public Task<XDocument> UpdateReportAsync(string url, string type, Encoding encoding, bool enableUpdateUser, string data, string uploadFilePath) {
        throw new NotImplementedException();
    }

    public Task<(bool, byte[])> GetReportFileAsync(string url, string fileType, int reportId) {
        throw new NotImplementedException();
    }

    public Task<XDocument> AutoGenerateAsync(string url, string type, Encoding encoding, int userMode, int labelMode, string data, string fileName, bool calculateEnable = true) {
        throw new NotImplementedException();
    }

    public Task<XElement> GetDefinitionListAsync(string url, string reportName, int lebelID) {
        throw new NotImplementedException();
    }

    public Task<XElement> GetDefinitionListAsync(string url, string reportName) {
        throw new NotImplementedException();
    }

    public Task<XElement> GetDefinitionListAsync(string url, int lebelID) {
        throw new NotImplementedException();
    }

    public Task<XElement> GetDefinitionListAsync(string url) {
        throw new NotImplementedException();
    }

    public Task<XElement> GetDefinitionDetailAsync(string url, string topId) {
        throw new NotImplementedException();
    }
}
