//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

namespace AutoReportLib.Models; 
/// <summary>
/// RP_入力帳票取得ヘッダ
/// </summary>
public class 入力帳票取得ヘッダ {
    public int ID { get; }
    public string? 処理済 { get; }
    public DateTime 処理日時 { get; }

    public bool Is処理済 => 処理済 is not null && 処理済 == 処理済Val;

    public const string 処理済Val = "1";
    public const string 処理未Val = "0";
}
