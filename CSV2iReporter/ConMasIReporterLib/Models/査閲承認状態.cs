using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConMasIReporterLib.Models;

public static class 編集ステータス {
    public const string 入力前 = "0";
    public const string 編集中 = "1";
    public const string 承認待ち = "2";
    public const string 差し戻し = "3";
    public const string 入力完了 = "4";
}
/// <summary>
/// 承認クラスター入力値
/// </summary>
public static class 承認入力値 {
    /// <summary>
    /// 未入力：ブランク 又は 0
    /// </summary>
    public const string 未入力 = "0";
    public const string 承認待ち = "2";
    public const string 承認 = "4";
}

/// <summary>
/// 査閲クラスター入力値
/// </summary>
public static class 査閲入力値 {
    public const string 査閲 = "4";
    /// <summary>
    /// 未入力：ブランク 又は 0
    /// </summary>
    public const string 未入力 = "0";
}
