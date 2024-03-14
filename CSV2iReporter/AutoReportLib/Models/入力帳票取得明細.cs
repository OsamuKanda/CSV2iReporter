//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

namespace AutoReportLib.Models.自動起票DB; 
/// <summary>
/// RP_入力帳票取得明細
/// </summary>
public class 入力帳票取得明細 {
    public int ID { get; }
    public int 元帳票ID { get; }
    public int 帳票ID { get; }
    public int シートNO { get; }
    public int クラスターID { get; }
    public string? データ { get; }
    public string? カラム名 { get; }
}
