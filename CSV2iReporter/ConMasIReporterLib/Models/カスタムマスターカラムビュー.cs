namespace AutoReportMakeLib.Models.i_ReporterDB; 
/// <summary>
/// view_mst_custom_column
/// </summary>
public class カスタムマスターカラムビュー {
    /// <summary>カスタムマスターID</summary>
    public int? master_id { get; }
    /// <summary>カスタムマスターキー</summary>
    public string? master_key { get; }
    /// <summary>カスタムマスター名称</summary>
    public string? master_name { get; }
    /// <summary> </summary>
    public int? master_type { get; }
    /// <summary>タブレット保存可否</summary>
    public int? mobile_save { get; }
    /// <summary>有効期限</summary>
    public DateTime? expire { get; }
    /// <summary>手動ダウンロード</summary>
    public int? download_type { get; }
    /// <summary>タブレット保持期間</summary>
    public int? life_time { get; }
    /// <summary>カスタムマスター表示順</summary>
    public int? master_display_number { get; }
    /// <summary>レコードキーカラム名称</summary>
    public string? record_key_name { get; }
    /// <summary>レコードバリューカラム名称</summary>
    public string? record_value_name { get; }
    /// <summary>備考</summary>
    public string? remark { get; }
    /// <summary>レコードバージョン</summary>
    public string? record_version { get; }
    /// <summary>カスタムカラムフィールド番号</summary>
    public int? field_no { get; }
    /// <summary>カスタムカラムフィールド名称</summary>
    public string? value { get; }
    /// <summary>カスタムフィールド型</summary>
    public string? type { get; }
    /// <summary>登録端末</summary>
    public string? sys_regist_term { get; }
    /// <summary>登録者</summary>
    public string? sys_regist_user { get; }
    /// <summary>登録日時</summary>
    public DateTime? sys_regist_time { get; }
    /// <summary>更新端末</summary>
    public string? sys_update_term { get; }
    /// <summary>更新者</summary>
    public string? sys_update_user { get; }
    /// <summary>更新日時</summary>
    public DateTime? sys_update_time { get; }
}
