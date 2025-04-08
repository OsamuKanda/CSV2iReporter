using ConMasIReporterLib;
using Serilog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.CompilerServices;
using System.Runtime.Intrinsics.Arm;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.Serialization;
using static AutoReportLib.Models.AutoReport;


namespace CSV2iReporter.Models;


public class AutoReport() {
    /// <summary>
    /// ラベル設定クラス
    /// </summary>
    public class Label {
        /// <summary>
        /// アイコンを示す接頭辞
        /// </summary>
        private const string prefix = "[icon=";
        /// <summary>
        /// アイコンを示す接尾辞
        /// </summary>
        private const string suffix = "]";

        public Label() : this("") { }
        /// <summary>
        /// ラベル指定コンストラクタ（アイコンとラベルを同時指定）
        /// </summary>
        /// <param name="label">アイコンは1～10で指定し、ラベルは/で括る</param>
        public Label(string label) {
            var sp = label.IndexOf(prefix, StringComparison.CurrentCulture);
            if (sp != -1) {
                var len = label[(sp)..].IndexOf(suffix, StringComparison.CurrentCulture);
                if (len != -1) {
                    //アイコン指定がある場合はアイコンを指定して残りをラベルにする
                    icon = int.Parse(label[(sp + prefix.Length)..len]);
                    labelName = label[(sp + len + suffix.Length)..];
                    return;
                }
            }
            //アイコン指定が無い場合はラネルをそのまま指定
            labelName = label;
        }
        /// <summary>
        /// アイコン番号１～１０
        /// </summary>
        [XmlElement]
        public int? icon;
        /// <summary>
        /// ラベル名
        /// </summary>
        [XmlElement("name")]
        public string labelName;
    }
    /// <summary>
    /// 備考インターフェース
    /// </summary>
    public interface IRemarks {
#pragma warning disable IDE1006 // 命名スタイル
        public string? remarksValue1 { get; set; }
        public string? remarksValue2 { get; set; }
        public string? remarksValue3 { get; set; }
        public string? remarksValue4 { get; set; }
        public string? remarksValue5 { get; set; }
        public string? remarksValue6 { get; set; }
        public string? remarksValue7 { get; set; }
        public string? remarksValue8 { get; set; }
        public string? remarksValue9 { get; set; }
        public string? remarksValue10 { get; set; }
#pragma warning restore IDE1006 // 命名スタイル
    }
    /// <summary>
    /// 備考を配列化するクラス
    /// </summary>
    public class Remarks() {
        private readonly IRemarks? _rem;
        public Remarks(AutoReport.IRemarks rem) : this() {
            _rem = rem;
        }
        public string? this[int index] {
            get {
                return index switch {
                    0 => _rem?.remarksValue1,
                    1 => _rem?.remarksValue2,
                    2 => _rem?.remarksValue3,
                    3 => _rem?.remarksValue4,
                    4 => _rem?.remarksValue5,
                    5 => _rem?.remarksValue6,
                    6 => _rem?.remarksValue7,
                    7 => _rem?.remarksValue8,
                    8 => _rem?.remarksValue9,
                    9 => _rem?.remarksValue10,
                    _ => null,
                };
            }
            set {
                if( _rem is null){
                    return;
                }
                switch (index) {
                    case 0: _rem.remarksValue1 = value; break;
                    case 1: _rem.remarksValue2 = value; break;
                    case 2: _rem.remarksValue3 = value; break;
                    case 3: _rem.remarksValue4 = value; break;
                    case 4: _rem.remarksValue5 = value; break;
                    case 5: _rem.remarksValue6 = value; break;
                    case 6: _rem.remarksValue7 = value; break;
                    case 7: _rem.remarksValue8 = value; break;
                    case 8: _rem.remarksValue9 = value; break;
                    case 9: _rem.remarksValue10 = value; break;
                }
            }
        }
    }
    /// <summary>
    /// 備考インターフェース
    /// </summary>
    public interface ISystemKey {
#pragma warning disable IDE1006 // 命名スタイル
        public string? systemKey1 { get; set; }
        public string? systemKey2 { get; set; }
        public string? systemKey3 { get; set; }
        public string? systemKey4 { get; set; }
        public string? systemKey5 { get; set; }
#pragma warning restore IDE1006 // 命名スタイル
    }
    /// <summary>
    /// 備考を配列化するクラス
    /// </summary>
    public class SystemKey() {
        private readonly ISystemKey? _sys;
        public SystemKey(AutoReport.ISystemKey sys): this() {
            _sys = sys;
        }
        public string? this[int index] {
            get {
                return index switch {
                    0 => _sys?.systemKey1,
                    1 => _sys?.systemKey2,
                    2 => _sys?.systemKey3,
                    3 => _sys?.systemKey4,
                    4 => _sys?.systemKey5,
                    _ => null,
                };
            }
            set {
                if (_sys is null) {
                    return;
                }
                switch (index) {
                    case 0: _sys.systemKey1 = value; break;
                    case 1: _sys.systemKey2 = value; break;
                    case 2: _sys.systemKey3 = value; break;
                    case 3: _sys.systemKey4 = value; break;
                    case 4: _sys.systemKey5 = value; break;
                }
            }
        }
    }
//    public interface IReference {
//        public string? referenceType1 { get; set; }
//        public string? referenceName1 { get; set; }
//        public string? referenceValue1 { get; set; }
//        public string? referenceType2 { get; set; }
//        public string? referenceName2 { get; set; }
//        public string? referenceValue2 { get; set; }
//        public string? referenceType3 { get; set; }
//        public string? referenceName3 { get; set; }
//        public string? referenceValue3 { get; set; }
//        public string? referenceType4 { get; set; }
//        public string? referenceName4 { get; set; }
//        public string? referenceValue4 { get; set; }
//        public string? referenceType5 { get; set; }
//        public string? referenceName5 { get; set; }
//        public string? referenceValue5 { get; set; }
//        public string? referenceType6 { get; set; }
//        public string? referenceName6 { get; set; }
//        public string? referenceValue6 { get; set; }
//        public string? referenceType7 { get; set; }
//        public string? referenceName7 { get; set; }
//        public string? referenceValue7 { get; set; }
//        public string? referenceType8 { get; set; }
//        public string? referenceName8 { get; set; }
//        public string? referenceValue8 { get; set; }
//        public string? referenceType9 { get; set; }
//        public string? referenceName9 { get; set; }
//        public string? referenceValue9 { get; set; }
//        public string? referenceType10 { get; set; }
//        public string? referenceName10 { get; set; }
//        public string? referenceValue10 { get; set; }
//    }
//    public class Reference(References r, int index) {
//        private readonly References _r = r;
//        private readonly int _index = index;
//        public Reference this[int Index] {
//            string? Type {
//            get {
//                return _index switch {
//                    0 => _r._ref.referenceType1,
//                    1 => _r._ref.referenceType2,
//                    2 => _r._ref.referenceType3,
//                    3 => _r._ref.referenceType4,
//                    4 => _r._ref.referenceType5,
//                    5 => _r._ref.referenceType6,
//                    6 => _r._ref.referenceType7,
//                    7 => _r._ref.referenceType8,
//                    8 => _r._ref.referenceType9,
//                    9 => _r._ref.referenceType10,
//                    _ => null,
//                };
//            }
//            set {
//                switch (index) {
//                    case 0: _r._ref.referenceType1 = value; break;
//                    case 1: _r._ref.referenceType2 = value; break;
//                    case 2: _r._ref.referenceType3 = value; break;
//                    case 3: _r._ref.referenceType4 = value; break;
//                    case 4: _r._ref.referenceType5 = value; break;
//                    case 5: _r._ref.referenceType6 = value; break;
//                    case 6: _r._ref.referenceType7 = value; break;
//                    case 7: _r._ref.referenceType8 = value; break;
//                    case 8: _r._ref.referenceType9 = value; break;
//                    case 9: _r._ref.referenceType10 = value; break;
//                };
//            }
//        }
//        string? Name {
//            get {
//                return _index switch {
//                    0 => _r._ref.referenceName1,
//                    1 => _r._ref.referenceName2,
//                    2 => _r._ref.referenceName3,
//                    3 => _r._ref.referenceName4,
//                    4 => _r._ref.referenceName5,
//                    5 => _r._ref.referenceName6,
//                    6 => _r._ref.referenceName7,
//                    7 => _r._ref.referenceName8,
//                    8 => _r._ref.referenceName9,
//                    9 => _r._ref.referenceName10,
//                    _ => null,
//                };
//            }
//            set {
//                switch (index) {
//                    case 0: _r._ref.referenceName1 = value; break;
//                    case 1: _r._ref.referenceName2 = value; break;
//                    case 2: _r._ref.referenceName3 = value; break;
//                    case 3: _r._ref.referenceName4 = value; break;
//                    case 4: _r._ref.referenceName5 = value; break;
//                    case 5: _r._ref.referenceName6 = value; break;
//                    case 6: _r._ref.referenceName7 = value; break;
//                    case 7: _r._ref.referenceName8 = value; break;
//                    case 8: _r._ref.referenceName9 = value; break;
//                    case 9: _r._ref.referenceName10 = value; break;
//                };
//            }
//        }
//        string? Value {
//            get {
//                return _index switch {
//                    0 => _r._ref.referenceValue1,
//                    1 => _r._ref.referenceValue2,
//                    2 => _r._ref.referenceValue3,
//                    3 => _r._ref.referenceValue4,
//                    4 => _r._ref.referenceValue5,
//                    5 => _r._ref.referenceValue6,
//                    6 => _r._ref.referenceValue7,
//                    7 => _r._ref.referenceValue8,
//                    8 => _r._ref.referenceValue9,
//                    9 => _r._ref.referenceValue10,
//                    _ => null,
//                };
//            }
//            set {
//                switch (index) {
//                    case 0: _r._ref.referenceValue1 = value; break;
//                    case 1: _r._ref.referenceValue2 = value; break;
//                    case 2: _r._ref.referenceValue3 = value; break;
//                    case 3: _r._ref.referenceValue4 = value; break;
//                    case 4: _r._ref.referenceValue5 = value; break;
//                    case 5: _r._ref.referenceValue6 = value; break;
//                    case 6: _r._ref.referenceValue7 = value; break;
//                    case 7: _r._ref.referenceValue8 = value; break;
//                    case 8: _r._ref.referenceValue9 = value; break;
//                    case 9: _r._ref.referenceValue10 = value; break;
//                };
//            }
//        }
//    }
//    }
//}
//    /// <summary>
//    /// 参照インターフェース
//    /// </summary>
//    public class References(AutoReport.IReference r) {
//        public IReference _ref = r;
//        private Reference _reference[];
//        public References this[int index] {
//            get {
//                return _reference[index];
//            }
//            set {
//                _reference[index] = value;
//            }
//        }

//    }

    public class Top : IRemarks, ISystemKey {
        public Top() {
            remarks = new(this);
            systemKey = new(this);
        }
        public Top(int DefTopID ) : this() {
            defTopId = DefTopID;
        }
        /// <summary>
        /// 帳票定義ID
        /// </summary>
        [XmlElement]
        public int defTopId = 0;
        /// <summary>
        /// 帳票定義名（APIドキュメントには記載がないが、これが無いと動かない）
        /// </summary>
        [XmlElement]
        public string defTopName = "";
        /// <summary>
        /// 帳票名称（必須）
        /// </summary>
        [XmlElement]
        public string repTopName = "";
        /// <summary>
        /// 帳票備考１
        /// </summary>
        [XmlElement, Browsable(false)]
        public string? remarksValue1 { get; set; }
        /// <summary>
        /// 帳票備考２
        /// </summary>
        [XmlElement, Browsable(false)]
        public string? remarksValue2 { get; set; }
        /// <summary>
        /// 帳票備考３
        /// </summary>
        [XmlElement, Browsable(false)]
        public string? remarksValue3 { get; set; }
        /// <summary>
        /// 帳票備考４
        /// </summary>
        [XmlElement, Browsable(false)]
        public string? remarksValue4 { get; set; }
        /// <summary>
        /// 帳票備考５
        /// </summary>
        [XmlElement, Browsable(false)]
        public string? remarksValue5 { get; set; }
        /// <summary>
        /// 帳票備考６
        /// </summary>
        [XmlElement, Browsable(false)]
        public string? remarksValue6 { get; set; }
        /// <summary>
        /// 帳票備考７
        /// </summary>
        [XmlElement, Browsable(false)]
        public string? remarksValue7 { get; set; }
        /// <summary>
        /// 帳票備考８
        /// </summary>
        [XmlElement, Browsable(false)]
        public string? remarksValue8 { get; set; }
        /// <summary>
        /// 帳票備考９
        /// </summary>
        [XmlElement, Browsable(false)]
        public string? remarksValue9 { get; set; }
        /// <summary>
        /// 帳票備考１０
        /// </summary>
        [XmlElement, Browsable(false)]
        public string? remarksValue10 { get; set; }
        [XmlIgnore]
        public Remarks remarks;

        /// <summary>
        /// 作成モード 
        /// </summary>
        public enum CreateRoleMode : uint {
            /// <summary>
            /// すべて
            /// </summary>
            All = 0,
            /// <summary>
            /// 作成ユーザーが所属するグループのみ
            /// </summary>
            GroupOnly = 1,
            /// <summary>
            /// 作成ユーザーが所属するグループと上位のみ
            /// </summary>
            GroupHigh = 2,
            /// <summary>
            /// 作成ユーザーが所属するグループと上位下位
            /// </summary>
            GroupHighLow = 3

        };
        /// <summary>
        /// 作成モード 
        /// </summary>
        [XmlElement]
        public CreateRoleMode? createRoleMode;
        [XmlIgnore, Browsable(false)]
#pragma warning disable IDE1006 // 命名スタイル
        public bool createRoleModeSpecified => createRoleMode.HasValue;
#pragma warning restore IDE1006 // 命名スタイル
        /// <summary>
        /// 作成ユーザーID
        /// </summary>
        [XmlElement]
        public string? createUserId;
        /// <summary>
        /// PDF差し替えページ指定 0:全シート対象 1～n：ページ指定　（カンマ区切りで複数ページ指定）
        /// </summary>
        [XmlElement]
        public string? pdfReplacePage;
        /// <summary>
        /// PDFファイル名
        /// </summary>
        [XmlElement]
        public string? pdfFileName;
        /// <summary>
        /// システムキー1
        /// </summary>
        [XmlElement, Browsable(false)]
        public string? systemKey1 { get; set; }
        /// <summary>
        /// システムキー2
        /// </summary>
        [XmlElement, Browsable(false)]
        public string? systemKey2 { get; set; }
        /// <summary>
        /// システムキー3
        /// </summary>
        [XmlElement, Browsable(false)]
        public string? systemKey3 { get; set; }
        /// <summary>
        /// システムキー4
        /// </summary>
        [XmlElement, Browsable(false)]
        public string? systemKey4 { get; set; }
        /// <summary>
        /// システムキー5
        /// </summary>
        [XmlElement, Browsable(false)]
        public string? systemKey5 { get; set; }
        /// <summary>
        /// 配列化システムキー
        /// </summary>
        public SystemKey systemKey;

        // シートとクラスタ
        public class Sheet: IRemarks {
            /// <summary>
            /// 指定なしコンストラクタ
            /// </summary>
            public Sheet() {
                this.remarks = new(this);
            }
            /// <summary>
            /// シート番号指定コンストラクタ
            /// </summary>
            /// <param name="sheetNo">シート番号1～</param>
            public Sheet(int sheetNo) : this() {
                // 配下のシート番号をすべて上書きする
                this.sheetNo = sheetNo;
                foreach (var c in this.cluster) {
                    c.sheetNo = sheetNo;
                }
            }
            /// <summary>
            /// シート番号
            /// </summary>
            [XmlElement]
            public int sheetNo = 0;
            /// <summary>
            /// 帳票定義名（指定不要）
            /// </summary>
            [XmlElement]
            public string defSheetName = "";    // API仕様書に記載漏れてます
            /// <summary>
            /// シート名（指定不要）
            /// </summary>
            [XmlElement]
            public string? sheetName;
            /// <summary>
            /// シート備考1
            /// </summary>
            [XmlElement, Browsable(false)]
            public string? remarksValue1 { get; set; }
            /// <summary>
            /// シート備考2
            /// </summary>
            [XmlElement, Browsable(false)]
            public string? remarksValue2 { get; set; }
            /// <summary>
            /// シート備考3
            /// </summary>
            [XmlElement, Browsable(false)]
            public string? remarksValue3 { get; set; }
            /// <summary>
            /// シート備考4
            /// </summary>
            [XmlElement, Browsable(false)]
            public string? remarksValue4 { get; set; }
            /// <summary>
            /// シート備考5
            /// </summary>
            [XmlElement, Browsable(false)]
            public string? remarksValue5 { get; set; }
            /// <summary>
            /// シート備考6
            /// </summary>
            [XmlElement, Browsable(false)]
            public string? remarksValue6 { get; set; }
            /// <summary>
            /// シート備考7
            /// </summary>
            [XmlElement, Browsable(false)]
            public string? remarksValue7 { get; set; }
            /// <summary>
            /// シート備考8
            /// </summary>
            [XmlElement, Browsable(false)]
            public string? remarksValue8 { get; set; }
            /// <summary>
            /// シート備考9
            /// </summary>
            [XmlElement, Browsable(false)]
            public string? remarksValue9 { get; set; }
            /// <summary>
            /// シート備考10
            /// </summary>
            [XmlElement, Browsable(false)]
            public string? remarksValue10 { get; set; }
            /// <summary>
            /// 備考を配列要素に変換
            /// </summary>
            [XmlIgnore]
            public Remarks remarks;

            /// <summary>
            /// 参考資料種別1（LIB:登録済みのドキュメントIDから指定、ADDこのAPIで"D"行の追加するzipファイル名から指定）
            /// </summary>
            [XmlElement]
            public string? referenceType1;
            /// <summary>
            /// 参考資料名1
            /// </summary>
            [XmlElement]
            public string? referenceName1;
            /// <summary>
            /// 参考資料参照先文字列1（参考資料種別がLIBの場合ドキュメントID、参考資料種別がADDの場合Zipに同封のファイル名）
            /// </summary>
            [XmlElement]
            public string? referenceValue1;
            /// <summary>
            /// <summary>
            /// 参考資料種別2（LIB:登録済みのドキュメントIDから指定、ADDこのAPIで"D"行の追加するzipファイル名から指定）
            /// </summary>
            [XmlElement]
            public string? referenceType2;
            /// <summary>
            /// 参考資料名2
            /// </summary>
            [XmlElement]
            public string? referenceName2;
            /// <summary>
            /// 参考資料参照先文字列2（参考資料種別がLIBの場合ドキュメントID、参考資料種別がADDの場合Zipに同封のファイル名）
            /// </summary>
            [XmlElement]
            public string? referenceValue2;
            /// <summary>
            /// 参考資料種別3（LIB:登録済みのドキュメントIDから指定、ADDこのAPIで"D"行の追加するzipファイル名から指定）
            /// </summary>
            [XmlElement]
            public string? referenceType3;
            /// <summary>
            /// 参考資料名3
            /// </summary>
            [XmlElement]
            public string? referenceName3;
            /// <summary>
            /// 参考資料参照先文字列3（参考資料種別がLIBの場合ドキュメントID、参考資料種別がADDの場合Zipに同封のファイル名）
            /// </summary>
            [XmlElement]
            public string? referenceValue3;
            /// <summary>
            /// 参考資料種別4（LIB:登録済みのドキュメントIDから指定、ADDこのAPIで"D"行の追加するzipファイル名から指定）
            /// </summary>
            [XmlElement]
            public string? referenceType4;
            /// <summary>
            /// 参考資料名4
            /// </summary>
            [XmlElement]
            public string? referenceName4;
            /// <summary>
            /// 参考資料参照先文字列4（参考資料種別がLIBの場合ドキュメントID、参考資料種別がADDの場合Zipに同封のファイル名）
            /// </summary>
            [XmlElement]
            public string? referenceValue4;
            /// <summary>
            /// 参考資料種別%（LIB:登録済みのドキュメントIDから指定、ADDこのAPIで"D"行の追加するzipファイル名から指定）
            /// </summary>
            [XmlElement]
            public string? referenceType5;
            /// <summary>
            /// 参考資料名5
            /// </summary>
            [XmlElement]
            public string? referenceName5;
            /// <summary>
            /// 参考資料参照先文字列5（参考資料種別がLIBの場合ドキュメントID、参考資料種別がADDの場合Zipに同封のファイル名）
            /// </summary>
            [XmlElement]
            public string? referenceValue5;
            /// <summary>
            /// 参考資料種別6（LIB:登録済みのドキュメントIDから指定、ADDこのAPIで"D"行の追加するzipファイル名から指定）
            /// </summary>
            [XmlElement]
            public string? referenceType6;
            /// <summary>
            /// 参考資料名6
            /// </summary>
            [XmlElement]
            public string? referenceName6;
            /// <summary>
            /// 参考資料参照先文字列6（参考資料種別がLIBの場合ドキュメントID、参考資料種別がADDの場合Zipに同封のファイル名）
            /// </summary>
            [XmlElement]
            public string? referenceValue6;
            /// <summary>
            /// 参考資料種別7（LIB:登録済みのドキュメントIDから指定、ADDこのAPIで"D"行の追加するzipファイル名から指定）
            /// </summary>
            [XmlElement]
            public string? referenceType7;
            /// <summary>
            /// 参考資料名7
            /// </summary>
            [XmlElement]
            public string? referenceName7;
            /// <summary>
            /// 参考資料参照先文字列7（参考資料種別がLIBの場合ドキュメントID、参考資料種別がADDの場合Zipに同封のファイル名）
            /// </summary>
            [XmlElement]
            public string? referenceValue7;
            /// <summary>
            /// 参考資料種別8（LIB:登録済みのドキュメントIDから指定、ADDこのAPIで"D"行の追加するzipファイル名から指定）
            /// </summary>
            [XmlElement]
            public string? referenceType8;
            /// <summary>
            /// 参考資料名8
            /// </summary>
            [XmlElement]
            public string? referenceName8;
            /// <summary>
            /// 参考資料参照先文字列8（参考資料種別がLIBの場合ドキュメントID、参考資料種別がADDの場合Zipに同封のファイル名）
            /// </summary>
            [XmlElement]
            public string? referenceValue8;
            /// <summary>
            /// 参考資料種別9（LIB:登録済みのドキュメントIDから指定、ADDこのAPIで"D"行の追加するzipファイル名から指定）
            /// </summary>
            [XmlElement]
            public string? referenceType9;
            /// <summary>
            /// 参考資料名9
            /// </summary>
            [XmlElement]
            public string? referenceName9;
            /// <summary>
            /// 参考資料参照先文字列9（参考資料種別がLIBの場合ドキュメントID、参考資料種別がADDの場合Zipに同封のファイル名）
            /// </summary>
            [XmlElement]
            public string? referenceValue9;
            /// <summary>
            /// 参考資料種別10（LIB:登録済みのドキュメントIDから指定、ADDこのAPIで"D"行の追加するzipファイル名から指定）
            /// </summary>
            [XmlElement]
            public string? referenceType10;
            /// <summary>
            /// 参考資料名10
            /// </summary>
            [XmlElement]
            public string? referenceName10;
            /// <summary>
            /// 参考資料参照先文字列10（参考資料種別がLIBの場合ドキュメントID、参考資料種別がADDの場合Zipに同封のファイル名）
            /// </summary>
            [XmlElement]
            public string? referenceValue10;

            public class Cluster : IRemarks {
                public Cluster() {
                    this.remarks = new(this);
                }
                public Cluster(int sheetNo, int clusterId, string value) : this() {
                    this.sheetNo = sheetNo;         // 構造体の親にもsheetNoがあるのに。。。APIの仕様XMLがそうなってる
                    this.clusterId = clusterId;
                    this.value = value;
                }
                /// <summary>
                /// 
                /// </summary>
                [XmlElement]
                public int sheetNo;
                [XmlElement]
                public int clusterId;
                [XmlElement, Browsable(false)]
                public string? clusterName;
                /// <summary>
                /// クラスタへの設定値。記入不要マーク設定時は"{verified}"を設定
                /// </summary>
                [XmlElement]
                public string value = "";
                /// <summary>
                /// クラスター備考1
                /// </summary>
                [XmlElement]
                public string? remarksValue1 { get; set; }
                /// <summary>
                /// クラスター備考2
                /// </summary>
                [XmlElement]
                public string? remarksValue2 { get; set; }
                /// <summary>
                /// クラスター備考3
                /// </summary>
                [XmlElement]
                public string? remarksValue3 { get; set; }
                /// <summary>
                /// クラスター備考4
                /// </summary>
                [XmlElement]
                public string? remarksValue4 { get; set; }
                /// <summary>
                /// クラスター備考5
                /// </summary>
                [XmlElement]
                public string? remarksValue5 { get; set; }
                /// <summary>
                /// クラスター備考6
                /// </summary>
                [XmlElement]
                public string? remarksValue6 { get; set; }
                /// <summary>
                /// クラスター備考7
                /// </summary>
                [XmlElement]
                public string? remarksValue7 { get; set; }
                /// <summary>
                /// クラスター備考8
                /// </summary>
                [XmlElement]
                public string? remarksValue8 { get; set; }
                /// <summary>
                /// クラスター備考9
                /// </summary>
                [XmlElement]
                public string? remarksValue9 { get; set; }
                /// <summary>
                /// クラスター備考10
                /// </summary>
                [XmlElement]
                public string? remarksValue10 { get; set; }
                /// <summary>
                /// 備考を配列要素に変換
                /// </summary>
                [XmlIgnore]
                public Remarks remarks;
                /// <summary>
                /// コメント入力値 （チェッククラスター、トグル選択クラスターのみ適用）
                /// </summary>
                [XmlElement]
                public string? comment;
                /// <summary>
                /// クラスター名称を定義設定から書き変えます（指定時はコピー不可）
                /// </summary>
                [XmlElement]
                public string? clusterNameChange;
                /// <summary>
                /// Phoneリスト表示順（指定時はコピー不可）
                /// </summary>
                [XmlElement]
                public string? mobileListDisplayNo;
                /// <summary>
                /// iPhoneリスト表示、非表示（指定時はコピー不可）
                /// </summary>
                public enum MobileDisplay : int {
                    /// <summary>
                    /// 表示しない
                    /// </summary>
                    NoDisplay = 0,
                    /// <summary>
                    /// 表示する
                    /// </summary>
                    Display = 1
                }
                /// <summary>
                /// iPhoneリスト表示、非表示（指定時はコピー不可）
                /// </summary>
                [XmlElement]
                public int? mobileDisplay;
                [XmlIgnore, Browsable(false)]
#pragma warning disable IDE1006 // 命名スタイル
                public bool mobileDisplaySpecified => mobileDisplay.HasValue;
#pragma warning restore IDE1006 // 命名スタイル
                /// <summary>
                /// クラスタに指定する最大値
                /// </summary>
                [XmlElement]
                public string? maximum;
                /// <summary>
                /// クラスタに指定する最小値
                /// </summary>
                [XmlElement]
                public string? minimum;
                /// <summary>
                /// クラスタに指定する最大閾値
                /// </summary>
                [XmlElement]
                public string? allowMaxValue;
                /// <summary>
                /// クラスタに指定する最小閾値
                /// </summary>
                [XmlElement]
                public string? allowMinValue;
                /// <summary>
                ///  編集ユーザーID
                /// </summary>
                [XmlElement]
                public string? editUser;
                /// <summary>
                ///  編集ユーザー名 
                /// </summary>
                [XmlElement]
                public string? editUserName;
                /// <summary>
                ///  編集日時（ yyyy/MM/dd HH:mm:ss形式）
                /// </summary>
                [XmlElement]
                public string? editTime;
                /// <summary>
                /// 編集した端末のGPS座標
                /// </summary>
                public class Gps {
                    public Gps() { }
                    public Gps(decimal lat,decimal lon,decimal alt){
                        this.lat = lat;
                        this.lon = lon;
                        this.alt = alt;
                    }
                    public Gps(decimal lat, decimal lon) {
                        this.lat = lat;
                        this.lon = lon;
                    }
                    /// <summary>
                    /// 緯度
                    /// </summary>
                    [XmlElement]
                    public decimal lat;
                    /// <summary>
                    /// 経度
                    /// </summary>
                    [XmlElement]
                    public decimal lon;
                    /// <summary>
                    /// 高度
                    /// </summary>
                    [XmlElement]
                    public decimal alt;
                }
                [XmlElement]
                public Gps? gps;
            }
            /// <summary>
            /// 
            /// </summary>
            [XmlArray("clusters")]
            [XmlArrayItem("cluster")]
            public List<Cluster> cluster = [];

            /// <summary>
            /// クラスター呼び出し音声認識辞書
            /// </summary>
            public class GramCluster {
                public GramCluster() { }
                public GramCluster(int sheetNo,int clusterId) : this() {
                    this.sheetNo = sheetNo;
                    this.clusterId = clusterId;
                }
                //public int sheetNo;
                //public int clusterId;
                /// <summary>
                /// シート番号
                /// </summary>
                public int sheetNo = 0;
                /// <summary>
                /// クラスターID
                /// </summary>
                public int clusterId = 0;
                /// <summary>
                /// 呼び出し名称
                /// </summary>
                public string words = "";
                /// <summary>
                /// アンサーバック 
                /// </summary>
                public string answerBack = "";
                /// <summary>
                /// アンサーバック入力値読み上げ
                /// </summary>
                public enum InputAnswerBack : int {
                    /// <summary>
                    /// 読み上げしない
                    /// </summary>
                    Disable = 0,
                    /// <summary>
                    /// 読み上げする
                    /// </summary>
                    Enable = 1
                }
                public InputAnswerBack inputAnswerBack = InputAnswerBack.Disable;

            }
            /// <summary>
            /// クラスター呼び出し音声認識辞書
            /// </summary>
            [XmlArray("grammers")]
            [XmlArrayItem("cluster")]
            public List<GramCluster> gramCluster = [];
        }
        /// <summary>
        /// シート情報
        /// </summary>
        [XmlArray("sheets")]
        [XmlArrayItem("sheet")]
        public List<Sheet> sheet = [];

        /// <summary>
        /// 参考資料設定
        /// </summary>
        public class Document {
            /// <summary>
            /// 指定なしコンストラクタ
            /// </summary>
            public Document() { }
            /// <summary>
            /// ドキュメントファイル名又はURL（ファイルの場合は同一Zip内に圧縮１帳票内で、重複不可）
            /// </summary>
            [XmlElement]
            public string documentName = "";
            /// <summary>
            /// ドキュメントの表示名
            /// </summary>
            [XmlElement]
            public string documentDisplayName = "";
            public enum DocumentIcon : int {
                Pdf,
                Excel,
                PowerPoint,
                Word,
                Xvl,
                Movie,
                Photo,
                Sound,
                Others
            }
            /// <summary>
            /// ドキュメントアイコンの区分
            /// </summary>
            [XmlElement]
            public string? documentIcon;
            /// <summary>
            /// 保管区分
            /// </summary>
            public enum DocumentSaveType : int {
                FileSystem = 0,         // ファイルシステム
                Url = 1                 // URL
            }
            /// <summary>
            /// 保管区分
            /// </summary>
            [XmlElement]
            public DocumentSaveType? documentSaveType;
            [XmlIgnore, Browsable(false)]
#pragma warning disable IDE1006 // 命名スタイル
            public bool documentSaveTypeSpecified => documentSaveType.HasValue;
#pragma warning restore IDE1006 // 命名スタイル
            /// <summary>
            /// タブレット保存可否
            /// </summary>
            public enum DocumentMobileSave : int {
                Enable = 0,        // 可能
                Disable = 1,       // 不可能
            }
            /// <summary>
            /// タブレット保存可否
            /// </summary>
            [XmlElement]
            public DocumentMobileSave? documentMobileSave;
            [XmlIgnore, Browsable(false)]
#pragma warning disable IDE1006 // 命名スタイル
            public bool documentMobileSaveSpecified => documentMobileSave.HasValue;
#pragma warning restore IDE1006 // 命名スタイル
            /// <summary>
            /// 有効期限（yyyy/MM/dd）
            /// </summary>
            [XmlElement]
            public string? useEndTime = "";
            /// <summary>
            /// ラベル
            /// </summary>
            [XmlArray("addLabels")]
            [XmlArrayItem("label")]
            public List<Label> label = [];
            /// <summary>
            /// 参照可能権限(グループIDをセミコロンで区切って複数設定
            /// </summary>
            [XmlElement]
            public string? referRole;
        }
        /// <summary>
        /// 参照資料
        /// </summary>
        [XmlArray("documents")]
        [XmlArrayItem("document")]
        public List<Document> document = [];

        /// <summary>
        /// スケジュール設定
        /// </summary>
        public class Schedule {
            /// <summary>
            /// 引数無しコンストラクタ
            /// </summary>
            public Schedule() { }
            public Schedule(int taskId, string user) : this() {
                this.taskId = taskId;
                this.user = user;
            }
            [XmlElement]
            public int taskId = 0;
            [XmlElement]
            public string user = "";
            [XmlElement]
            public string startDate = DateTime.Now.ToString("yyyy/MM/dd");
            [XmlElement]
            public string? endDate;
            [XmlElement]
            public string? comment1;
            [XmlElement]
            public string? comment2;
            public enum AutoDownload : int {
                Disable = 0,
                Enable = 1
            }
            [XmlElement]
            public AutoDownload? autoDownload;
            [XmlIgnore, Browsable(false)]
#pragma warning disable IDE1006 // 命名スタイル
            public bool autoDownloadSpecified => autoDownload.HasValue;
#pragma warning restore IDE1006 // 命名スタイル
        }
        /// <summary>
        /// スケジュール設定
        /// </summary>
        [XmlArray("schedules")]
        [XmlArrayItem("schedule")]
        public List<Schedule> schedule = [];

        /// <summary>
        /// 通知メッセージ
        /// </summary>
        public class Notice() {
            /// <summary>
            /// 通知メッセージアイコン
            /// </summary>
            public enum Icon : int {
                Information = 0,    // 情報
                Note = 1,           // 注意
                Warning = 2         // 警告
            }
            /// <summary>
            /// 通知メッセージアイコン
            /// </summary>
            [XmlElement]
            public Icon? icon;
            [XmlIgnore, Browsable(false)]
#pragma warning disable IDE1006 // 命名スタイル
            public bool iconSpecified => icon.HasValue;
#pragma warning restore IDE1006 // 命名スタイル
            /// <summary>
            /// 通知メッセージヘッダ
            /// </summary>
            [XmlElement]
            public string? messageHeader;
            /// <summary>
            /// 通知メッセージボディ
            /// </summary>
            [XmlElement]
            public string? messageBody;
            /// <summary>
            /// 通知メッセージ画像ヘッダ
            /// </summary>
            [XmlElement]
            public string? imageHeader;
            /// <summary>
            /// 通知メッセージ参照URL
            /// </summary>
            [XmlElement]
            public string? url;
            /// <summary>
            /// 使用開始日
            /// </summary>
            [XmlElement]
            public string? useStartDate;
            /// <summary>
            /// 使用終了日
            /// </summary>
            [XmlElement]
            //public DateTime? useEndDate;
            public string? useEndDate;
            /// <summary>
            /// イメージファイル名
            /// </summary>
            [XmlElement]
            //public string? imageFileName;
            public string imageFileName = "";
        }
        /// <summary>
        /// 通知メッセージ
        /// </summary>
        [XmlElement]
        public Notice? notice;
        /// <summary>
        /// 追加ラベル情報
        /// </summary>
        [XmlArray("addLabels")]
        [XmlArrayItem("label")]
        public List<Label> label = [];
    }

    /// <summary>
    /// 自動帳票作成データ
    /// </summary>
    [XmlArray("conmas")]
    [XmlArrayItem("top")]
    public List<Top> top = [];
}
