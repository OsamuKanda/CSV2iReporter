using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;


namespace AutoReportLib.Models;

public class AutoReport {
    public class Conmas {
        public class Top {
            /// <summary>
            /// 帳票ID（必須）
            /// </summary>
            public int defTopID;
            /// <summary>
            /// 帳票名
            /// </summary>
            public string? repTopName;
            public string? remarksValue1;
            public string? remarksValue2;
            public string? remarksValue3;
            public string? remarksValue4;
            public string? remarksValue5;
            public string? remarksValue6;
            public string? remarksValue7;
            public string? remarksValue8;
            public string? remarksValue9;
            public string? remarksValue10;
            /// <summary>
            /// 作成モード 
            /// </summary>
            public enum CreateRoleMode : uint {
                All = 0,            // すべて
                GroupOnly = 1,      // 作成ユーザーが所属するＧのみ
                GroupHigh = 2,      // 作成ユーザーが所属するＧ＋上位のみ
                GroupHighLow = 3    // 作成ユーザーが所属するＧ＋上位下位のみ

            };
            public CreateRoleMode? createRoleMode;

            public int? createUserId;
            public int[]? pdfReplacePages;
            public string? pdfFileName;
            public int? systemkey1;
            public int? systemkey2;
            public int? systemkey3;
            public int? systemkey4;
            public int? systemkey5;
            public class Notice {
                public enum Icon : int {
                    Information = 0,    // 情報
                    Note = 1,           // 注意
                    Warning = 2         // 警告
                }
                public Icon? icon;

                public string? messageHeader;
                public string? messageBody;
                public string? imageHeader;
                public string? url;
                public DateTime? useStartDate;
                public DateTime? useEndDate;
                public string? imageFileName;
            }
            public Notice? notice;

            public class AddLebels {
                public class Label {
                    public int? icon;
                    public string? name;
                }
                public Label[]? label;

            }
            public AddLebels? addLabels;

            public class Documents {
                public class Document {
                    public string? documentName;
                    public string? documentDisplayName;
                    public string? documentIcon;
                    public enum DocumentSaveType : int {
                        FileSystem = 0,         // ファイルシステム
                        Url = 1                 // URL
                    }
                    public DocumentSaveType? documentSaveType;

                    public enum DocumentMobileSave : int {
                        Can = 0,        // 可能
                        Cant = 1,       // 不可能

                    }
                    public DocumentMobileSave? documentMobileSave;

                    public string? useEndTime;
                    public class AddLabels {
                        public Label[]? label;
                    }
                    public AddLabels? addLabels;

                    public string? referRole;
                }
                public Document[]? documents;

            }
            public Documents? document;

            public class Schedules {
                public class Schedule {
                    public int? taskId;
                    public string? user;
                    public string? startDate;
                    public string? endDate;
                    public string? comment1;
                    public string? comment2;
                    public enum AutoDownload : int {
                        Disable = 0,
                        Enable = 1
                    }
                    public AutoDownload? autoDownload;
                }
                public Schedule[]? schedule;
            }
            public Schedules? schedules;

            public class Sheets {
                public class Sheet {
                    public int? sheetNo;
                    public string? sheetName;
                    public string? remarksValue1;
                    public string? remarksValue2;
                    public string? remarksValue3;
                    public string? remarksValue4;
                    public string? remarksValue5;
                    public string? remarksValue6;
                    public string? remarksValue7;
                    public string? remarksValue8;
                    public string? remarksValue9;
                    public string? remarksValue10;
                    public string? referenceType1;
                    public string? referenceName1;
                    public string? referenceValue1;
                    public string? referenceType2;
                    public string? referenceName2;
                    public string? referenceValue2;
                    public string? referenceType3;
                    public string? referenceName3;
                    public string? referenceValue3;
                    public string? referenceType4;
                    public string? referenceName4;
                    public string? referenceValue4;
                    public string? referenceType5;
                    public string? referenceName5;
                    public string? referenceValue5;
                    public string? referenceType6;
                    public string? referenceName6;
                    public string? referenceValue6;
                    public string? referenceType7;
                    public string? referenceName7;
                    public string? referenceValue7;
                    public string? referenceType8;
                    public string? referenceName8;
                    public string? referenceValue8;
                    public string? referenceType9;
                    public string? referenceName9;
                    public string? referenceValue9;
                    public string? referenceType10;
                    public string? referenceName10;
                    public string? referenceValue10;
                    public class Clusters {
                        public class Cluster {
                            public int sheetNo;
                            public int clusterId;
                            public string? value;
                            public string? remarksvalue1;
                            public string? remarksValue2;
                            public string? remarksValue3;
                            public string? remarksValue4;
                            public string? remarksValue5;
                            public string? remarksValue6;
                            public string? remarksValue7;
                            public string? remarksValue8;
                            public string? remarksValue9;
                            public string? remarksValue10;
                            public string? comment;
                            public string? clusterNameChange;
                            public int? mobileListDisplayNo;
                            public int? mobileDisplay;
                            public string? maximum;
                            public string? minimum;
                            public string? allowMaxValue;
                            public string? allowMinValue;
                            public string? editUser;
                            public string? editUserName;
                            public string? editTime;        // yyyy/MM/dd HH:mm:ss形式
                            public class Gps {
                                public decimal lat;         // 緯度
                                public decimal lon;         // 経度
                                public decimal? alt;        // 高度
                            }
                            public Gps[]? gps;

                            public class Grammers {
                                public class Cluster {
                                    public int sheetNo;
                                    public int clusterId;
                                    public string? words;
                                    public string? answerBack;
                                    public string? inputAnswerBack;

                                }
                                public Cluster[]? clusters;

                            }
                            public Grammers? grammers;

                        }
                        public Cluster[]? cluster;

                    }
                    public Clusters? cluster;

                }
                public Sheet[]? sheet;

            }
            public Sheets? sheets;

        }
        public Top[]? top;

    }
    public Conmas conmas = new();
}
