using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ConMasIReporterLib.Models;

public class ReportResult {
    public string? topId { get; set; }
    public string? code { get; set; }
}

/// <summary>
/// ConMasWebAPIの結果XMLをパースするためのクラス
/// </summary>
public class ConMasWebAPIResult {
    public XDocument? ResultDocument { get; set; }

    public ConMasWebAPIResult(XDocument? resultDocument) {
        ResultDocument = resultDocument;
    }

    public string? ErrorCode =>
        ResultDocument?.Element("conmas")?.Element("error")?.Element("code")?.Value;

    public bool IsSuccess =>
        ResultDocument?.Element("conmas")?.Element("error")?.Element("code")?.Value == "0";

    public IEnumerable<ReportResult> ReportResults() {
        var results = ResultDocument?.Element("conmas")?.Element("results")?.Elements("result");
        if (results != null) {
            foreach (var result in results) {
                yield return new ReportResult {
                    topId = result.Element("topId")?.Value,
                    code = result.Element("code")?.Value
                };
            }
        }
    }
}
