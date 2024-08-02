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

        var fileTemplate = "| {Timestamp:HH:mm:ss.fff} | {Level:u4} | {ThreadId:00}:{ThreadName,24} | {Message:j} | {NewLine}{Exception}";
        var consoleTemplate = "| {Timestamp:HH:mm:ss.fff} | {Level:u4} | {ThreadId:00}:{ThreadName,24} | {Message:j} | {NewLine}{Exception}";
        var logConf = new LoggerConfiguration();
        logConf
           //.MinimumLevel.Debug()
           .Enrich.WithExceptionDetails()
           .Enrich.WithThreadId()
           .Enrich.WithThreadName().Enrich.WithProperty("ThreadName", "_")
           .Enrich.FromLogContext();
        if (!string.IsNullOrEmpty(logFolder)) {
            logConf.WriteTo.File( path:logFolder, rollingInterval: RollingInterval.Day, outputTemplate: fileTemplate);
        }
        //ログレベルの設定
        logConf.MinimumLevel.Is(
            configuration.GetSection("Log")["Level"] switch {
                "V" or "0" => LogEventLevel.Verbose,
                "D" or "1" => LogEventLevel.Debug,
                "I" or "2" => LogEventLevel.Information,
                "W" or "3" => LogEventLevel.Warning,
                "E" or "4" => LogEventLevel.Error,
                "F" or "5" => LogEventLevel.Fatal,
                _ => LogEventLevel.Verbose
            }
        );

        // デバッグコンソールと、コンソールに出力
        logConf
            .WriteTo.Debug(outputTemplate: consoleTemplate)
            .WriteTo.Console(outputTemplate: consoleTemplate);
        Log.Logger = logConf.CreateLogger();
    }
}
