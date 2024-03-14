using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Extensions.Configuration;
using Serilog;
using ConMasIReporterLib;
using ConMasIReporterLib.Models;

namespace AutoReportMake.Services; 
/// <summary>
/// ComMasWebAPI関連
/// </summary>
public class ComMasWebAPIService2 : ComMasWebAPIService {
    public ComMasWebAPIService2(IConfiguration sec) : base(
         sec["Address"] ?? "",
         sec["User"] ?? "",
         sec["Pass"] ?? "",
         sec["ProxyServer"] ?? null,
         sec["ProxyUserName"] ?? null,
         sec["ProxyPassword"] ?? null,
         sec["ProxyDomain"] ?? null
        ) {
    }
}
