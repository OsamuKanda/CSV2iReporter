using CSV2iReporter;


// 引数（オプション読み込み）
var settingFile = ".\\appsettings.json";
for(var i = 0; i < args.Length; i++) {
    if ((args[i].Equals("-paramfile", StringComparison.CurrentCultureIgnoreCase)) || (args[i].Equals("-pf", StringComparison.CurrentCultureIgnoreCase))) {
        if (i != args.Length - 1) {
            settingFile = args[i + 1];
        }
    }
}

//設定ファイル読み出し
var configuration = AutoReportLib.Init.ReadSettingFile(settingFile);

//Serilog初期化
AutoReportLib.Init.InitLog(configuration);

// i-Reporter自動帳票連携ソフト準備
var apl = new ReportMakeFromCSV(configuration);

// 設定ファイル、帳票定義ファイルチェック
var result = await apl.Init();

// 設定ファイル、帳票定義ファイルがOKの場合
if (result == true) {
    // 登録元ファイルより帳票生成
    await apl.Execute();
}


