//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

namespace AutoReportLib.Models.自動起票DB; 
/// <summary>
/// RP_帳票データヘッダ
/// </summary>
public class 帳票データヘッダ {
    public int ID { get; }
    public int 帳票ID { get; }
    public string? 処理済 { get; }
    public DateTime 処理日時 { get; }

    public bool Is処理済 => 処理済 != null && 処理済 == 処理済Val;

    public const string 処理済Val = "1";
    public const string 処理未Val = "0";

    public 帳票データヘッダ() {
    }

    public 帳票データヘッダ(int 帳票ID) {
        this.帳票ID = 帳票ID;
        処理済 = 処理未Val;
    }
}
