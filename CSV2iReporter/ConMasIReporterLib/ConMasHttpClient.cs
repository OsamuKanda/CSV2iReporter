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

    public string Address { get; }
    public string User { get; }
    public string Password { get; }

    public ConMasHttpClient(string address, string user, string password, string? proxyServer = null, string? proxyUserName = null, string? proxyPassword = null, string? proxyDomain = null) {
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
        Address = address;
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
    public async Task<bool> LoginAsync(string address, string userId, string password) {
        var resultXml = await PostAndXDocumentParseAsync(address, new Dictionary<string, string> {
            {"command", "Login" },
            {"user", userId },
            {"password", password }
        });
        var code = resultXml?.Element("conmas")?.Element("loginResult")?.Element("code")?.Value;
        return code == "0";
    }

    //インターフェイス側コメント参照
    public async Task<bool> LogoutAsync(string address) {
        var resultXml = await PostAndXDocumentParseAsync(address, new Dictionary<string, string> {
            {"command", "Logout" }
        });
        var code = resultXml?.Element("conmas")?.Element("error")?.Element("code")?.Value;

        return code == "0";
    }

    //インターフェイス側コメント参照
    public async Task<XDocument> AutoGenerateAsync(string address, string type, Encoding encoding, int userMode, int labelMode, string uploadFilePath, bool calculateEnable = true) {
        var urlString = $"{address}?command=AutoGenerate&type={type}&encoding={ToStrEncoding(encoding)}&userMode={userMode}&labelMode={labelMode}{(calculateEnable ? "&calculateEnable=1" : "")}";
        var content = new MultipartFormDataContent();
        var fileContent = new StreamContent(File.OpenRead(uploadFilePath));
        fileContent.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment") {
            FileName = Path.GetFileName(uploadFilePath)
        };
        content.Add(fileContent);

        return await PostAndXDocumentParseAsync(urlString, content);
    }

    //インターフェイス側コメント参照
    public async Task<XDocument> AutoGenerateAsync(string address, string type, Encoding encoding, int userMode, int labelMode, string data, string fileName, bool calculateEnable = true) {
        var urlString = $"{address}?command=AutoGenerate&type={type}&encoding={ToStrEncoding(encoding)}&userMode={userMode}&labelMode={labelMode}{(calculateEnable ? "&calculateEnable=1" : "")}";
        var content = new MultipartFormDataContent();
        var fileContent = new StreamContent(new MemoryStream(encoding.GetBytes(data)));
        fileContent.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment") {
            FileName = fileName
        };
        content.Add(fileContent);

        return await PostAndXDocumentParseAsync(urlString, content);
    }

    //インターフェイス側コメント参照
    public async Task<XElement> GetDefinitionListAsync(string address, string reportName = "", int lebelID = -9) {
        var content = new FormUrlEncodedContent(new Dictionary<string, string> {
            {"command", "GetDefinitionList" },
            {"labelId", lebelID.ToString() },
            {"word", reportName },
            {"wordTargetName", (reportName == "") ? "false": "true"}
        });

        var result = await httpClient.PostAsync(address, content);
        var resultXml = XElement.Parse(Encoding.UTF8.GetString(await result.Content.ReadAsByteArrayAsync()));

        return resultXml;
    }

    public Task<XElement> GetDefinitionListAsync(string address) {
        return GetDefinitionListAsync(address, "", -9);
    }

    public Task<XElement> GetDefinitionListAsync(string address,string reportName) {
        return GetDefinitionListAsync(address, reportName, -9);
    }

    public Task<XElement> GetDefinitionListAsync(string address, int lebelID) {
        return GetDefinitionListAsync(address, "", lebelID);
    }
    //インターフェイス側コメント参照
    public async Task<XElement> GetDefinitionDetailAsync(string address, string topId) {
        var content = new FormUrlEncodedContent(new Dictionary<string, string> {
            {"command", "GetDefinitionDetail" },
            {"topId", $"{topId}" }
        });

        var result = await httpClient.PostAsync(address, content);
        var resultXml = XElement.Parse(Encoding.UTF8.GetString(await result.Content.ReadAsByteArrayAsync()));

        return resultXml;
    }

    //インターフェイス側コメント参照
    public async Task<XDocument> UpdateReportAsync(string address, string type, Encoding encoding, bool enableUpdateUser, string data, string uploadFilePath) {
        var urlString = $"{address}?command=UpdateReport&type={type}&encoding={ToStrEncoding(encoding)}{(enableUpdateUser ? "&userMode=1" : "")}";
        var content = new MultipartFormDataContent();
        var fileContent = new StreamContent(new MemoryStream(encoding.GetBytes(data)));
        fileContent.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment") {
            FileName = uploadFilePath
        };
        content.Add(fileContent);

        return await PostAndXDocumentParseAsync(urlString, content);
    }

    //インターフェイス側コメント参照
    public async Task<XDocument> UpdateReportAsync(string address, string type, string encoding, bool enableUpdateUser, string uploadFilePath) {
        var urlString = $"{address}?command=UpdateReport&type={type}&encoding={encoding}";
        var content = new MultipartFormDataContent();
        var fileContent = new StreamContent(File.OpenRead(uploadFilePath));
        fileContent.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment") {
            FileName = Path.GetFileName(uploadFilePath)
        };
        content.Add(fileContent);

        return await PostAndXDocumentParseAsync(urlString, content);
    }

    private async Task<XDocument> PostAndXDocumentParseAsync(string address, IDictionary<string, string> param) =>
        await PostAndXDocumentParseAsync(address, new FormUrlEncodedContent(param));

    private async Task<XDocument> PostAndXDocumentParseAsync(string address, HttpContent? content) {
        var result = await httpClient.PostAsync(address, content);
        return XDocument.Parse(Encoding.UTF8.GetString(await result.Content.ReadAsByteArrayAsync()));
    }

    //インターフェイス側コメント参照
    public async Task<(bool, byte[])> GetReportFileAsync(string address, string fileType, int reportId) {
        var content = new FormUrlEncodedContent(new Dictionary<string, string> {
            {"command", "GetReportFile" },
            {"fileType", fileType },
            {"report", reportId.ToString() }
        });

        var result = await httpClient.PostAsync(address, content);
        //var contentType = result.Content.Headers.ContentType.ToString() ?? "";
        var arrayData = await result.Content.ReadAsByteArrayAsync();
        //var resultXml = XElement.Parse(Encoding.UTF8.GetString(await result.Content.ReadAsByteArrayAsync()));

        return (result.StatusCode.ToString() == "OK", arrayData);
    }
    private static string ToStrEncoding(Encoding encoding) =>
        encoding == Encoding.UTF8 ? "65001" : "932";
}
