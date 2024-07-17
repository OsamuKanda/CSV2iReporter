using Microsoft.Extensions.Configuration;
using ConMasIReporterLib;

namespace AutoReportMake.Services; 
/// <summary>
/// ComMasWebAPI関連
/// </summary>
public class ComMasWebAPIService2(IConfiguration sec) : ComMasWebAPIService(
     sec["Url"] ?? "",
     sec["User"] ?? "",
     sec["Password"] ?? "",
     sec["ProxyServer"] ?? null,
     sec["ProxyUserName"] ?? null,
     sec["ProxyPassword"] ?? null,
     sec["ProxyDomain"] ?? null
    ) {
}
