using Microsoft.Extensions.Configuration;
using Serilog;
using Serilog.Exceptions;
using System.Reflection;

namespace AutoReportLib;
public class Init() {

    public static string SettingFileFullPath { get; set; } = $@"{Directory.GetCurrentDirectory()}\appsettings.json";

    //設定ファイル読み出し
    public static IConfigurationRoot ReadSettingFile(string settingFileFullPath) {
        // .\\が指定されている場合はカレントフォルダに変換する
        SettingFileFullPath = settingFileFullPath.Replace(".\\", Directory.GetCurrentDirectory() + "\\");

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
        var logFolder = configuration.GetSection("Log")["Path"];
        var logLevel = configuration.GetSection("Log")["LogLevel"] switch {
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
            config.WriteTo.File(logFolder, rollingInterval: RollingInterval.Day, outputTemplate: template);
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
