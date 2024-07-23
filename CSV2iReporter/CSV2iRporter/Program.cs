using CSV2iReporter;


// 引数（オプション読み込み）
// EXEと同じフォルダ
var directoryName = Path.GetDirectoryName(Environment.ProcessPath);
// EXEと同じファイル名（拡張子無し）
var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(Environment.ProcessPath);
// 設定ファイルの指定
var settingFile = Path.Join(directoryName, fileNameWithoutExtension + ".json");

for (var i = 0; i < args.Length; i++) {
    if ((args[i].Equals("-parameterfile", StringComparison.CurrentCultureIgnoreCase)) || (args[i].Equals("-pf", StringComparison.CurrentCultureIgnoreCase))) {
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


