using Microsoft.Extensions.Configuration;
using Serilog;
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
        if(settingFileFullPath.Substring(0,2) == @".\") {
            SettingFileFullPath = Directory.GetCurrentDirectory() + @"\" + settingFileFullPath.Substring(2);
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
        var logLevel = configuration.GetSection("Log")["Level"] switch {
            "E" => 1,
            "W" => 2,
            "I" => 3,
            "N" => 0,
            _ => 3
        };
        var template = "| {Timestamp:HH:mm:ss.fff} | {Level:u4} | {ThreadId:00}:{ThreadName,24} | {Message:j} | {NewLine}{Exception}";
        var config = new LoggerConfiguration();
        config
           //.MinimumLevel.Debug()
           .Enrich.WithExceptionDetails()
           .Enrich.WithThreadId()
           .Enrich.WithThreadName().Enrich.WithProperty("ThreadName", "_")
           .Enrich.FromLogContext();
        if (!string.IsNullOrEmpty(logFolder)) {
            config.WriteTo.File( path:logFolder, rollingInterval: RollingInterval.Day, outputTemplate: template);
        }

        if (logLevel == 0) config.MinimumLevel.Fatal();
        if (logLevel == 1) config.MinimumLevel.Error();
        if (logLevel == 2) config.MinimumLevel.Warning();
        if (logLevel == 3) config.MinimumLevel.Information();

        config
            .WriteTo.Debug(outputTemplate: template)
            .WriteTo.Console();
        Log.Logger = config.CreateLogger();
    }
}
