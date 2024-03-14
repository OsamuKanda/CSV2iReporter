using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DbAccessLib;

public static class DbUtil {
    /// <summary>
    /// シングルクォーテーションでくくったカンマ区切りで返す。
    /// </summary>
    /// <param name="list"></param>
    /// <returns></returns>
    public static string CommaDelim(IEnumerable<string?> list) {
        var quatation = "'";
        return string.Join(",", list.Select(x => quatation + x + quatation));
    }
    public static string CommaDelim(IEnumerable<int> list) {
        return string.Join(",", list.Select(x => x.ToString()));
    }
    /// <summary>
    /// nullを除外して、シングルクォーテーションでくくったカンマ区切りで返す。
    /// </summary>
    /// <param name="list"></param>
    /// <returns></returns>
    public static string CommaDelimNotNull(IEnumerable<string?> list) =>
        CommaDelim(list.Where(m => m != null));
}
