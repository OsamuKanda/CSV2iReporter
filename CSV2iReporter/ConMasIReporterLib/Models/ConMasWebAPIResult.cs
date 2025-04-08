using System.Xml.Linq;

namespace ConMasIReporterLib.Models;

public class ReportResult {
    public string? TopId { get; set; }
    public string? Code { get; set; }
}

/// <summary>
/// ConMasWebAPIの結果XMLをパースするためのクラス
/// </summary>
public class ConMasWebAPIResult(XDocument? resultDocument) {
    public XDocument? ResultDocument { get; set; } = resultDocument;

    public string? ErrorCode =>
        ResultDocument?.Element("conmas")?.Element("error")?.Element("code")?.Value;

    public bool IsSuccess =>
        ResultDocument?.Element("conmas")?.Element("error")?.Element("code")?.Value == "0";

    public IEnumerable<ReportResult> ReportResults() {
        var results = ResultDocument?.Element("conmas")?.Element("results")?.Elements("result");
        if (results is not null) {
            foreach (var result in results) {
                yield return new ReportResult {
                    TopId = result.Element("topId")?.Value,
                    Code = result.Element("code")?.Value
                };
            }
        }
    }
}
