using ConMasIReporterLib;
using ConMasIReporterLib.Models;
using System.Text;

namespace AutoReportLib.Services; 
public static class ComMasWebAPIServiceEx {
    /// <summary>
    /// 帳票更新を呼び出し、
    /// 承認クラスターの状態を承認へ更新します。
    /// </summary>
    /// <param name="svc"></param>
    /// <param name="topId"></param>
    /// <param name="clusterId"></param>
    /// <param name="承認者ID"></param>
    /// <returns></returns>
    public static async Task 承認Async(this ComMasWebAPIService svc, int topId, int clusterId, string 承認者ID, int sheetNo = 1) {
        //承認
        var t = new 帳票更新CSVレイアウト.トップデータ(topId) {
            更新ユーザーID = 承認者ID,
            編集ステータス = 編集ステータス.入力完了
        };
        var s1 = new 帳票更新CSVレイアウト.シートデータ {
            シートＮＯ = $"{sheetNo}"
        };
        var c1 = new 帳票更新CSVレイアウト.クラスターデータ(clusterId) {
            クラスター入力値 = 承認入力値.承認,
            承認者ＩＤ = 承認者ID,
            承認日 = $"{DateTime.Now:yyyy/MM/dd}",
            承認者コメント = "承認コメント",
        };
        var sb = new StringBuilder()
            .Append(t.ToString())
            .Append(s1.ToString())
            .Append(c1.ToString())
            ;

        await svc.帳票更新Async(sb);
    }
}
