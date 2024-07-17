using System.Text;
using System.Xml.Linq;

namespace ConMasIReporterLib;

public interface IConMasWebAPIClient {
    /// <summary>
    /// ログイン
    /// command=Login
    /// </summary>
    /// <param name="url">i-Reporterサーバーアドレス</param>
    /// <param name="userId"></param>
    /// <param name="password"></param>
    /// <returns></returns>
    Task<bool> LoginAsync(string url, string userId, string password);

    /// <summary>
    /// ログアウト
    /// command=Logout
    /// </summary>
    /// <param name="url">i-Reporterサーバーアドレス</param>
    /// <returns>正常時はtrue</returns>
    Task<bool> LogoutAsync(string url);

    /// <summary>
    /// 自動帳票作成
    /// command=AutoGenerate
    /// </summary>
    /// <param name="url">i-Reporterサーバーアドレス</param>
    /// <param name="type"></param>
    /// <param name="encoding"></param>
    /// <param name="userMode"></param>
    /// <param name="labelMode"></param>
    /// <param name="uploadFilePath"></param>
    /// <returns></returns>
    Task<XDocument> AutoGenerateAsync(string url, string type, Encoding encoding, int userMode, int labelMode, string uploadFilePath, bool calculateEnable = true);

    /// <summary>
    /// 自動帳票作成
    /// command=AutoGenerate
    /// </summary>
    /// <param name="url">i-Reporterサーバーアドレス</param>
    /// <param name="type"></param>
    /// <param name="encoding"></param>
    /// <param name="userMode"></param>
    /// <param name="labelMode"></param>
    /// <param name="data"></param>
    /// <param name="fileName"></param>
    /// <returns></returns>
    Task<XDocument> AutoGenerateAsync(string url, string type, Encoding encoding, int userMode, int labelMode, string data, string fileName, bool calculateEnable = true);

    /// <summary>
    /// 帳票更新
    /// command=UpdateReport
    /// </summary>
    /// <param name="url">i-Reporterサーバーアドレス</param>
    /// <param name="type"></param>
    /// <param name="encoding"></param>
    /// <param name="enableUpdateUser">true  : アップロードファイル中の「更新ユーザー」のユーザーが帳票更新者として登録される。
    ///                                false : ※"更新ユーザー"がなければ、API実行ユーザーが帳票更新者
    /// </param>  
    /// <param name="uploadFilePath"></param>
    /// <returns></returns>
    Task<XDocument> UpdateReportAsync(string url, string type, string encoding, bool enableUpdateUser, string uploadFilePath);

    /// <summary>
    /// 帳票更新
    /// command=UpdateReport
    /// </summary>
    /// <param name="url">i-Reporterサーバーアドレス</param>
    /// <param name="type"></param>
    /// <param name="encoding"></param>
    /// <param name="enableUpdateUser">true  : アップロードファイル中の「更新ユーザー」のユーザーが帳票更新者として登録される。
    ///                                false : ※"更新ユーザー"がなければ、API実行ユーザーが帳票更新者
    /// <param name="data"></param>
    /// <param name="uploadFilePath"></param>
    /// <returns></returns>
    Task<XDocument> UpdateReportAsync(string url, string type, Encoding encoding, bool enableUpdateUser, string data, string uploadFilePath);

    /// <summary>
    /// 帳票ファイル取得
    /// command=GetReportFile
    /// </summary>
    /// <param name="url">i-Reporterサーバーアドレス</param>
    /// <param name="fileType"></param>
    /// <param name="reportId"></param>
    /// <returns></returns>
    Task<(bool, byte[])> GetReportFileAsync(string url, string fileType, int reportId);

    /// <summary>
    /// 定義一覧取得
    /// command=GetDefinitionList
    /// </summary>
    /// <param name="url"></param>
    /// <param name="labelID"></param>
    /// <param name="reportName"></param>
    /// <returns></returns>
    Task<XElement> GetDefinitionListAsync(string url);
    Task<XElement> GetDefinitionListAsync(string url, string reportName);
    Task<XElement> GetDefinitionListAsync(string url, int labelID);
    Task<XElement> GetDefinitionListAsync(string url, string reportName, int labelID);

    /// <summary>
    /// 定義簡易詳細情報取得
    /// </summary>
    /// <param name="url"></param>
    /// <param name="topId"></param>
    /// <returns></returns>
    Task<XElement> GetDefinitionDetailAsync(string url, string topId);
}
