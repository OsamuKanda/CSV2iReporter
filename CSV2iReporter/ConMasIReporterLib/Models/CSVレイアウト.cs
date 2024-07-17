using System.Text;

namespace ConMasIReporterLib.Models;

/// <summary>
/// CSVレイアウト（アップロード用） 作成用拡張クラス
/// </summary>
public static class CSVレイアウト {
    public static StringBuilder BeginH() =>
        new StringBuilder()
            .Append($"\"H\",\"defTopId\"");

    public static StringBuilder AddTopName(this StringBuilder sb) =>
        sb.Append($",\"repTopName\"");

    public static StringBuilder AddTOP備考情報(this StringBuilder sb, int index) =>
        sb.Append($",\"remarksValue{index}\"");

    public static StringBuilder addLabels(this StringBuilder sb) =>
        sb.Append($",\"addLabels\"");

    public static StringBuilder Addクラスター(this StringBuilder sb, int sheetNo, int clusterNo) =>
        sb.Append($",\"S{sheetNo}C{clusterNo}\"");

    public static StringBuilder Addクラスター(this StringBuilder sb, int clusterNo) =>
        sb.Addクラスター(1, clusterNo);

    public static StringBuilder Addクラスター(this StringBuilder sb, int sheetNo, IEnumerable<int> clusterNos) {
        foreach (var clusterNo in clusterNos) {
            sb.Addクラスター(sheetNo, clusterNo);
        }
        return sb;
    }

    public static StringBuilder Add備考(this StringBuilder sb, int remarksNo) =>
        sb.Append($",\"remarksValue{remarksNo}\"");

    public static StringBuilder AddSystemKey(this StringBuilder sb, int systemKeyNo) =>
        sb.Append($",\"systemKey{systemKeyNo}\"");

    public static StringBuilder Addクラスター(this StringBuilder sb, IEnumerable<int> clusterNos) =>
        sb.Addクラスター(1, clusterNos);

    public static StringBuilder End(this StringBuilder sb) =>
        sb.AppendLine();

    public static StringBuilder BeginR(this StringBuilder sb, int defTopId) =>
        sb.Append($"\"R\",\"{defTopId}\"");

    public static StringBuilder AddValue(this StringBuilder sb, object value) =>
        sb.Append($",\"{value}\"");
}
