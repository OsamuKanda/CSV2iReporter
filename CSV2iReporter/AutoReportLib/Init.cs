using Microsoft.Extensions.Configuration;
using Serilog;
using Serilog.Events;
using Serilog.Exceptions;
using System.Reflection;

namespace AutoReportLib;
public class Init() {

    public static string SettingFileFullPath { get; set; } = Path.Join(Path.GetDirectoryName(Environment.ProcessPath), Path.GetFileNameWithoutExtension(Environment.ProcessPath) + @".json");

    //設定ファイル読み出し
    public static IConfigurationRoot ReadSettingFile(string settingFileFullPath) {
        // ▼2024.07.16 Update TC Kanda (先頭以外に.\\を使った場合におかしな事になる）
        //// .\\が指定されている場合はカレントフォルダに変換する
        //SettingFileFullPath = settingFileFullPath.Replace(".\\", Directory.GetCurrentDirectory() + "\\");
        if (settingFileFullPath[0..2] == @".\") {
            SettingFileFullPath = Directory.GetCurrentDirectory() + @"\" + settingFileFullPath[2..];
        }
        // ▲2024.07.16 Update TC Kanda (先頭以外に.\\を使った場合におかしな事になる）

        //Console.WriteLine(System.IO.Directory.GetCurrentDirectory());
        //if (!File.Exists($@"{SettingFileFullPath}")) {

        //    file = "appsettings.json";
        //}

        return new ConfigurationBuilder()
          //.SetBasePath($@"{filepath}")
          .SetBasePath($@"{Path.GetDirectoryName(SettingFileFullPath)}").AddJsonFile(Path.GetFileName(SettingFileFullPath), true, true).Build();
    }

    //Serilog初期化
    public static void InitLog(IConfigurationRoot configuration) {
        var logFolder = configuration.GetSection("Log")["OutputPath"];
        //var logLevel = configuration.GetSection("Log")["Level"] switch {
        //    "V" => LogEventLevel.Verbose,
        //    "D" => LogEventLevel.Debug,
        //    "I" => LogEventLevel.Information,
        //    "W" => LogEventLevel.Warning,
        //    "E" => LogEventLevel.Error,
        //    "F" => LogEventLevel.Fatal,
        //    "0" => LogEventLevel.Verbose,
        //    "1" => LogEventLevel.Debug,
        //    "2" => LogEventLevel.Information,
        //    "3" => LogEventLevel.Warning,
        //    "4" => LogEventLevel.Error,
        //    "5" => LogEventLevel.Fatal,
        //    _ => LogEventLevel.Verbose
        //};
        var fileTemplate = "| {Timestamp:HH:mm:ss.fff} | {Level:u4} | {ThreadId:00}:{ThreadName,24} | {Message:j} | {NewLine}{Exception}";
        var consoleTemplate = "| {Timestamp:HH:mm:ss.fff} | {Level:u4} | {ThreadId:00}:{ThreadName,24} | {Message:j} | {NewLine}{Exception}";
        var config = new LoggerConfiguration();
        config
           //.MinimumLevel.Debug()
           .Enrich.WithExceptionDetails()
           .Enrich.WithThreadId()
           .Enrich.WithThreadName().Enrich.WithProperty("ThreadName", "_")
           .Enrich.FromLogContext();
        if (!string.IsNullOrEmpty(logFolder)) {
            config.WriteTo.File( path:logFolder, rollingInterval: RollingInterval.Day, outputTemplate: fileTemplate);
        }
        // エラーレベルの設定
        switch (configuration.GetSection("Log")["Level"]??"i".ToLower()) {
            // 出力無し
            case "v" or "0":
                config.MinimumLevel.Verbose();
                break;
            // デバッグ以上を出力
            case "d" or "1":
                config.MinimumLevel.Debug();
                break;
            // 情報以上を出力
            case "i" or "2":
                config.MinimumLevel.Information();
                break;
            // 警報以上を出力
            case "w" or "3":
                config.MinimumLevel.Warning();
                break;
            // エラー以上を出力
            case "e" or "4":
                config.MinimumLevel.Error();
                break;
            // 致命的エラーを出力
            case "f" or "5":
                config.MinimumLevel.Fatal();
                break;
            // 未指定の場合は情報以上を出力
            default:
                config.MinimumLevel.Information();
                break;
        }
        // デバッグコンソールと、コンソールに出力
        config
            .WriteTo.Debug(outputTemplate: consoleTemplate)
            .WriteTo.Console(outputTemplate: consoleTemplate);
        Log.Logger = config.CreateLogger();
    }
}
