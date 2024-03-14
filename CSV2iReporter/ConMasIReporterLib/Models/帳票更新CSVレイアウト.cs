using System.Text;

namespace ConMasIReporterLib.Models;

public class 帳票更新CSVレイアウト {
    public readonly StringBuilder sb = new StringBuilder();

    /// <summary>
    /// 標準的なトップデータを生成し
    /// StringBuilderとして返す(単独シートかつ、シートNO = 1)
    /// </summary>
    public static StringBuilder Createトップデータ(int topId, string ユーザーID, int sheetNo = 1) =>
        new StringBuilder()
            .Append(new 帳票更新CSVレイアウト.トップデータ(topId) {
                更新ユーザーID = ユーザーID,
                編集ステータス = 編集ステータス.承認待ち
            })
            .Append(new 帳票更新CSVレイアウト.シートデータ { シートＮＯ = $"{sheetNo}" });

    public class トップデータ {

        //NO 項目 説明 必須 タイプ 備考
        public string レコード区分 = "T";     // 1 | "T"固定 ◎ 文字列
        public int 帳票ＩＤ { get; }              // 2 | 対象となる帳票ＩＤ ◎ 数値 未入力時はエラー
        public string 定義名称 = "{ignore}";  // 3 | 定義名称 文字列 CSV手入力作成時の判別用のため、変更しても使用されない
        public string 帳票名称 = "{ignore}";  // 4 | 帳票名称 文字列 未入力は無視＝{ignore
        public string? ＴＯＰ備考情報１ = "{ignore}";      // 5 | 備考情報１ 文字列
        public string? ＴＯＰ備考情報２ = "{ignore}";      // 6 | 備考情報２ 文字列
        public string? ＴＯＰ備考情報３ = "{ignore}";      // 7 | 備考情報３ 文字列
        public string? ＴＯＰ備考情報４ = "{ignore}";      // 8 | 備考情報４ 文字列
        public string? ＴＯＰ備考情報５ = "{ignore}";      // 9 | 備考情報５ 文字列
        public string? ＴＯＰ備考情報６ = "{ignore}";      //10 | 備考情報６ 文字列
        public string? ＴＯＰ備考情報７ = "{ignore}";      //11 | 備考情報７ 文字列
        public string? ＴＯＰ備考情報８ = "{ignore}";      //12 備考情報８ 文字列
        public string? ＴＯＰ備考情報９ = "{ignore}";      //13 備考情報９ 文字列
        public string? ＴＯＰ備考情報１０ = "{ignore}";    //14 備考情報１０ 文字列
        public string? 編集ステータス = "{ignore}";        //15 | 編集ステータス 文字列
                                                    //空文字、０～４以外の場合は、無視
                                                    //０：入力前
                                                    //１：編集中
                                                    //２：承認待ち
                                                    //３：差し戻し
                                                    //４：入力完了
        public string? ラベル = "{ignore}";                //16 PDF差し替えページ指定 ０：全シート対象
                                                        // "[icon=1(~10)]"でアイコン指定（省略可）
                                                        //"/"区切りでラベル階層
                                                        //";"区切りで複数ラベル
                                                        //例）
                                                        //[icon=2] Label1/Label1-1;Label2
                                                        //(アイコンを省略した場合はアイコン１）
        public string? ラベルモード = "{ignore}";          //17 ラベルモード
                                                     //1:対象の帳票に付けられているラベルをすべて外す。
                                                     //1:以外は、なにもしない
        public string? PDF差し替えページ指定 = "";//18 | 
        public string? 差し替えPDFファイル名称 = "";//19 PDFファイル名
        public string? システムキー1 = "{ignore}";//20
        public string? システムキー2 = "{ignore}";//21
        public string? システムキー3 = "{ignore}";//22
        public string? システムキー4 = "{ignore}";//23
        public string? システムキー5 = "{ignore}";//24
        public string? 更新ユーザーID = "{ignore}";//25

        public トップデータ(int topId) {
            帳票ＩＤ = topId;
        }

        public override string ToString() =>
            new StringBuilder()
                .Append($"{レコード区分}")
                .Append($",{帳票ＩＤ}")
                .Append($",{定義名称}")
                .Append($",{帳票名称}")
                .Append($",{ＴＯＰ備考情報１}")
                .Append($",{ＴＯＰ備考情報２}")
                .Append($",{ＴＯＰ備考情報３}")
                .Append($",{ＴＯＰ備考情報４}")
                .Append($",{ＴＯＰ備考情報５}")
                .Append($",{ＴＯＰ備考情報６}")
                .Append($",{ＴＯＰ備考情報７}")
                .Append($",{ＴＯＰ備考情報８}")
                .Append($",{ＴＯＰ備考情報９}")
                .Append($",{ＴＯＰ備考情報１０}")
                .Append($",{編集ステータス}")
                .Append($",{ラベル}")
                .Append($",{ラベルモード}")
                .Append($",{PDF差し替えページ指定}")
                .Append($",{差し替えPDFファイル名称}")
                .Append($",{システムキー1}")
                .Append($",{システムキー2}")
                .Append($",{システムキー3}")
                .Append($",{システムキー4}")
                .Append($",{システムキー5}")
                .Append($",{更新ユーザーID}")
                .AppendLine()
                .ToString();
    }

    public class シートデータ {
        public class Cシート参考資料 {
            public string 種別 { get; set; } = "{ignore}";
            public string 名称 { get; set; } = "{ignore}";
            public string 参照先文字列 { get; set; } = "{ignore}";
        }

        //NO 項目 説明 必須 タイプ 備考
        public string レコード区分 = "S";// 固定 ◎ 文字列
        public string シートＮＯ { get; set; } = "1";// 帳票内のページ番号 ◎ 数値
        public string シート定義名称 { get; set; } = "{ignore}";// 定義名称 文字列 CSV手入力作成時の判別用のため、変更しても使用されない
        public string シート帳票名称 { get; set; } = "{ignore}";//帳票名称 文字列 無ければ定義の名称を使用

        public string[] シート備考情報 { get; } = Enumerable.Range(0, 10).Select(m => "{ignore}").ToArray();
        public Cシート参考資料[] シート参考資料 { get; } = Enumerable.Range(0, 10).Select(m => new Cシート参考資料()).ToArray();

        public override string ToString() {
            var sb = new StringBuilder()
                .Append($"{レコード区分}")
                .Append($",{シートＮＯ}")
                .Append($",{シート定義名称}")
                .Append($",{シート帳票名称}");
            foreach (var r in シート備考情報) {
                sb.Append($",{r}");
            }
            foreach (var r in シート参考資料) {
                sb.Append($",{r.種別}");
                sb.Append($",{r.名称}");
                sb.Append($",{r.参照先文字列}");
            }
            sb.AppendLine();
            return sb.ToString();
        }
    }

    public class クラスターデータ {
        public string レコード区分 = "C";// 1 レコード区分 "C"固定 ◎ 文字列
        public int クラスターＩＤ { get; }// 2 シート内クラスター番号 ◎ 数値
        public string クラスター名称 = "{ignore}";// 3 クラスター名称 文字列 CSV手入力作成時の判別用のため、変更しても使用されない
        public string クラスター入力値 = "{ignore}";// 4 クラスターに入力されたデータ 文字列 不正な値の場合は、無視　※参照
        public string クラスター備考情報１ = "{ignore}";// 5 備考情報１ 文字列
        public string クラスター備考情報２ = "{ignore}";// 6 備考情報２ 文字列
        public string クラスター備考情報３ = "{ignore}";// 7 備考情報３ 文字列
        public string クラスター備考情報４ = "{ignore}";// 8 備考情報４ 文字列
        public string クラスター備考情報５ = "{ignore}";// 9 備考情報５ 文字列
        public string クラスター備考情報６ = "{ignore}";//10 備考情報６ 文字列
        public string クラスター備考情報７ = "{ignore}";//11 備考情報７ 文字列
        public string クラスター備考情報８ = "{ignore}";//12 備考情報８ 文字列
        public string クラスター備考情報９ = "{ignore}";//13 備考情報９ 文字列
        public string クラスター備考情報１０ = "{ignore}";//14 備考情報１０ 文字列
        public string 申請者ＩＤ = "{ignore}";           //15 申請者ＩＤ 文字列
        public string 申請日 = "{ignore}";             //16 申請日 文字列
        public string 申請者コメント = "{ignore}";//17 申請者コメント 文字列
        public string 承認者ＩＤ = "{ignore}"; //18 承認者ＩＤ 文字列
        public string 承認日 = "{ignore}"; //19 承認日 文字列
        public string 承認者コメント = "{ignore}";//20 承認者コメント 文字列
        public string 印影イメージ = "{ignore}"; //21 印影イメージ 文字列
        public string コメント = "{ignore}"; //22 コメント 文字列 チェッククラスター、トグル選択のみ適用
        public string 編集ユーザーID = "{ignore}";//23 編集したユーザーのID
        public string 編集ユーザー名 = "{ignore}"; //24 編集したユーザーの名称
        public string 編集日時 = "{ignore}"; //25 編集した日時
        public string GPS緯度 = "{ignore}"; //26編集した端末のGPS座標の緯度
        public string GPS経度 = "{ignore}"; //27編集した端末のGPS座標の経度
        public string GPS高度 = "{ignore}"; //28編集した端末のGPS座標の高度

        public クラスターデータ(int クラスターＩＤ) {
            this.クラスターＩＤ = クラスターＩＤ;
        }

        public override string ToString() =>
            new StringBuilder()
                .Append($"{レコード区分}")
                .Append($",{クラスターＩＤ}")
                .Append($",{クラスター名称}")
                .Append($",{クラスター入力値}")
                .Append($",{クラスター備考情報１}")
                .Append($",{クラスター備考情報２}")
                .Append($",{クラスター備考情報３}")
                .Append($",{クラスター備考情報４}")
                .Append($",{クラスター備考情報５}")
                .Append($",{クラスター備考情報６}")
                .Append($",{クラスター備考情報７}")
                .Append($",{クラスター備考情報８}")
                .Append($",{クラスター備考情報９}")
                .Append($",{クラスター備考情報１０}")
                .Append($",{申請者ＩＤ}")
                .Append($",{申請日}")
                .Append($",{申請者コメント}")
                .Append($",{承認者ＩＤ}")
                .Append($",{承認日}")
                .Append($",{承認者コメント}")
                .Append($",{印影イメージ}")
                .Append($",{コメント}")
                .Append($",{編集ユーザーID}")
                .Append($",{編集ユーザー名}")
                .Append($",{編集日時}")
                .Append($",{GPS緯度}")
                .Append($",{GPS経度}")
                .Append($",{GPS高度}")
                .AppendLine()
                .ToString()
                ;
    }
}
