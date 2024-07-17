using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Xml.Linq;

namespace ConMasIReporterLib;

/// <summary>
/// 拡張WebClientクラス（Cookie対応）
/// <para>
/// Cookieによるセッション管理に対応したWebClienクラス
/// </para>
/// </summary>
/// <remarks>
/// コンストラクタ
/// </remarks>
public class HttpClientEx(HttpClientHandler hander) : HttpClient(hander) {
}

public class ConMasHttpClient : IDisposable, IConMasWebAPIClient {
    /// <summary>
    /// Http Client
    /// </summary>
    private readonly HttpClientEx httpClient;

    public string Url { get; }
    public string User { get; }
    public string Password { get; }

    public ConMasHttpClient(string url, string user, string password, string? proxyServer = null, string? proxyUserName = null, string? proxyPassword = null, string? proxyDomain = null) {
        HttpClientHandler ch;
        // プロキシ設定がある場合はプロキシ設定を行う
        if (string.IsNullOrEmpty(proxyServer)) {
            ch = new HttpClientHandler { UseProxy = false };
        } else {
            ch = new HttpClientHandler { UseProxy = true };
            if (proxyServer != "") {
                ch.Proxy = new WebProxy(proxyServer);
                if (proxyUserName != null) {
                    if (proxyDomain != null) {
                        ch.Proxy.Credentials = new NetworkCredential(proxyUserName, proxyPassword, proxyDomain);
                    } else {
                        ch.Proxy.Credentials = new NetworkCredential(proxyUserName, proxyPassword);
                    }
                }
            } else {
                ch.Proxy = WebRequest.GetSystemWebProxy();
            }
        }
        httpClient = new HttpClientEx(ch);
        Url = url;
        User = user;
        Password = password;
    }

    /// <summary>
    /// デストラクタ
    /// </summary>
    public void Dispose() {
        httpClient.Dispose();
        GC.SuppressFinalize(this);
    }

    //インターフェイス側コメント参照
    public async Task<bool> LoginAsync(string url, string userId, string password) {
        var resultXml = await PostAndXDocumentParseAsync(url, new Dictionary<string, string> {
            {"command", "Login" },
            {"user", userId },
            {"password", password }
        });
        var code = resultXml?.Element("conmas")?.Element("loginResult")?.Element("code")?.Value;
        return code == "0";
    }

    //インターフェイス側コメント参照
    public async Task<bool> LogoutAsync(string url) {
        var resultXml = await PostAndXDocumentParseAsync(url, new Dictionary<string, string> {
            {"command", "Logout" }
        });
        var code = resultXml?.Element("conmas")?.Element("error")?.Element("code")?.Value;

        return code == "0";
    }

    //インターフェイス側コメント参照
    public async Task<XDocument> AutoGenerateAsync(string url, string type, Encoding encoding, int userMode, int labelMode, string uploadFilePath, bool calculateEnable = true) {
        var urlString = $"{url}?command=AutoGenerate&type={type}&encoding={ToStrEncoding(encoding)}&userMode={userMode}&labelMode={labelMode}{(calculateEnable ? "&calculateEnable=1" : "")}";
        var content = new MultipartFormDataContent();
        var fileContent = new StreamContent(File.OpenRead(uploadFilePath));
        fileContent.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment") {
            FileName = Path.GetFileName(uploadFilePath)
        };
        content.Add(fileContent);

        return await PostAndXDocumentParseAsync(urlString, content);
    }

    //インターフェイス側コメント参照
    public async Task<XDocument> AutoGenerateAsync(string url, string type, Encoding encoding, int userMode, int labelMode, string data, string fileName, bool calculateEnable = true) {
        var urlString = $"{url}?command=AutoGenerate&type={type}&encoding={ToStrEncoding(encoding)}&userMode={userMode}&labelMode={labelMode}{(calculateEnable ? "&calculateEnable=1" : "")}";
        var content = new MultipartFormDataContent();
        var fileContent = new StreamContent(new MemoryStream(encoding.GetBytes(data)));
        fileContent.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment") {
            FileName = fileName
        };
        content.Add(fileContent);

        return await PostAndXDocumentParseAsync(urlString, content);
    }

    //インターフェイス側コメント参照
    public async Task<XElement> GetDefinitionListAsync(string url, string reportName = "", int lebelID = -9) {
        var content = new FormUrlEncodedContent(new Dictionary<string, string> {
            {"command", "GetDefinitionList" },
            {"labelId", lebelID.ToString() },
            {"word", reportName },
            {"wordTargetName", (reportName == "") ? "false": "true"}
        });

        var result = await httpClient.PostAsync(url, content);
        var resultXml = XElement.Parse(Encoding.UTF8.GetString(await result.Content.ReadAsByteArrayAsync()));

        return resultXml;
    }

    public Task<XElement> GetDefinitionListAsync(string url) {
        return GetDefinitionListAsync(url, "", -9);
    }

    public Task<XElement> GetDefinitionListAsync(string url,string reportName) {
        return GetDefinitionListAsync(url, reportName, -9);
    }

    public Task<XElement> GetDefinitionListAsync(string url, int lebelID) {
        return GetDefinitionListAsync(url, "", lebelID);
    }
    //インターフェイス側コメント参照
    public async Task<XElement> GetDefinitionDetailAsync(string url, string topId) {
        var content = new FormUrlEncodedContent(new Dictionary<string, string> {
            {"command", "GetDefinitionDetail" },
            {"topId", $"{topId}" }
        });

        var result = await httpClient.PostAsync(url, content);
        var resultXml = XElement.Parse(Encoding.UTF8.GetString(await result.Content.ReadAsByteArrayAsync()));

        return resultXml;
    }

    //インターフェイス側コメント参照
    public async Task<XDocument> UpdateReportAsync(string url, string type, Encoding encoding, bool enableUpdateUser, string data, string uploadFilePath) {
        var urlString = $"{url}?command=UpdateReport&type={type}&encoding={ToStrEncoding(encoding)}{(enableUpdateUser ? "&userMode=1" : "")}";
        var content = new MultipartFormDataContent();
        var fileContent = new StreamContent(new MemoryStream(encoding.GetBytes(data)));
        fileContent.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment") {
            FileName = uploadFilePath
        };
        content.Add(fileContent);

        return await PostAndXDocumentParseAsync(urlString, content);
    }

    //インターフェイス側コメント参照
    public async Task<XDocument> UpdateReportAsync(string url, string type, string encoding, bool enableUpdateUser, string uploadFilePath) {
        var urlString = $"{url}?command=UpdateReport&type={type}&encoding={encoding}";
        var content = new MultipartFormDataContent();
        var fileContent = new StreamContent(File.OpenRead(uploadFilePath));
        fileContent.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment") {
            FileName = Path.GetFileName(uploadFilePath)
        };
        content.Add(fileContent);

        return await PostAndXDocumentParseAsync(urlString, content);
    }

    private async Task<XDocument> PostAndXDocumentParseAsync(string url, IDictionary<string, string> param) =>
        await PostAndXDocumentParseAsync(url, new FormUrlEncodedContent(param));

    private async Task<XDocument> PostAndXDocumentParseAsync(string url, HttpContent? content) {
        var result = await httpClient.PostAsync(url, content);
        return XDocument.Parse(Encoding.UTF8.GetString(await result.Content.ReadAsByteArrayAsync()));
    }

    //インターフェイス側コメント参照
    public async Task<(bool, byte[])> GetReportFileAsync(string url, string fileType, int reportId) {
        var content = new FormUrlEncodedContent(new Dictionary<string, string> {
            {"command", "GetReportFile" },
            {"fileType", fileType },
            {"report", reportId.ToString() }
        });

        var result = await httpClient.PostAsync(url, content);
        //var contentType = result.Content.Headers.ContentType.ToString() ?? "";
        var arrayData = await result.Content.ReadAsByteArrayAsync();
        //var resultXml = XElement.Parse(Encoding.UTF8.GetString(await result.Content.ReadAsByteArrayAsync()));

        return (result.StatusCode.ToString() == "OK", arrayData);
    }
    private static string ToStrEncoding(Encoding encoding) =>
        encoding == Encoding.UTF8 ? "65001" : "932";
}
