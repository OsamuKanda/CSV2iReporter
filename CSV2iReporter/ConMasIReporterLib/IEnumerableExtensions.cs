using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/// <summary>
/// IEnumerable拡張メソッドクラス
/// </summary>
public static class IEnumerableExtensions {
    /// <summary>
    /// インデックス付コレクション(リスト)を返す
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="item"><c>IEnumerable<T></c></param>
    /// <returns>インデックス付コレクション</returns>
    public static IEnumerable<(T, int)> Select<T>(this IEnumerable<T> item) {
        return item.Select((v, i) => (v, i));
    }
}
