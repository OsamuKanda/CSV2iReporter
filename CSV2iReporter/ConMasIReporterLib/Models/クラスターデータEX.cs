using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConMasIReporterLib.Models;

public static class クラスターデータEX {
    public static StringBuilder クラスター入力値(this StringBuilder sb, int id, object value) {
        sb.Append(new 帳票更新CSVレイアウト.クラスターデータ(id) {
            クラスター入力値 = value.ToString() ?? ""
        });
        return sb;
    }

    public static StringBuilder クラスター入力値リスト(this StringBuilder sb, int offset, IEnumerable<object> list) {
        foreach (var (value, index) in list.Select()) {
            sb.クラスター入力値(offset + index, value);
        }
        return sb;
    }

    /// <summary>
    /// 査閲クラスターを査閲状態へ更新するデータを追加します
    /// </summary>
    /// <param name="sb">追加対象</param>
    /// <param name="id">対象クラスターID</param>
    /// <param name="ユーザーID"></param>
    /// <param name="査閲コメント"></param>
    /// <param name="日時"></param>
    /// <returns></returns>
    public static StringBuilder クラスター入力値_査閲(this StringBuilder sb, int id, string ユーザーID, string? 査閲コメント = null, DateTime? 日時 = null) {
        //査閲
        sb.Append(new 帳票更新CSVレイアウト.クラスターデータ(id) {
            クラスター入力値 = 査閲入力値.査閲,
            承認者ＩＤ = ユーザーID,
            承認日 = $"{(日時 ?? DateTime.Now):yyyy/MM/dd}",
            承認者コメント = 査閲コメント ?? "査閲コメント",
        });
        return sb;
    }

    /// <summary>
    /// 承認クラスターを承認待ちへ更新するデータを追加します
    /// </summary>
    /// <param name="sb">追加対象</param>
    /// <param name="id">対象クラスターID</param>
    /// <param name="ユーザーID"></param>
    /// <param name="コメント"></param>
    /// <param name="日時"></param>
    /// <returns></returns>
    public static StringBuilder クラスター入力値_承認申請(this StringBuilder sb, int id, string ユーザーID, string? コメント = null, DateTime? 日時 = null) {
        sb.Append(new 帳票更新CSVレイアウト.クラスターデータ(id) {
            クラスター入力値 = 承認入力値.承認待ち,
            申請者ＩＤ = ユーザーID,
            申請日 = $"{(日時 ?? DateTime.Now):yyyy/MM/dd}",
            申請者コメント = コメント ?? "申請コメント",
        });
        return sb;
    }
}
