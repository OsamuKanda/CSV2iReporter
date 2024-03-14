using AutoReportLib.Interfaces;
using AutoReportLib.Models.自動起票DB;
using DbAccessLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReportLib.Services;

public class AutoReportDbService {
    /// <summary>
    /// パラメータクエリ プレフィックス
    /// </summary>
    protected const string P = "@";

    protected readonly IMyTransactionFactory _transactionFactory = new MyTransactionFactory();//TODO DIで渡すようにする。

    public async Task<IMyTransaction> CreateAndBeginTransactionAsync() =>
        await _transactionFactory.CreateInstanceAsync();

    /// <summary>
    /// RP_入力帳票取得ヘッダテーブルのレコード取得
    /// </summary>
    public async Task<IList<入力帳票取得ヘッダ>> Get入力帳票取得ヘッダListAsync() =>
        await new StringBuilder()
            .AppendLine($"SELECT * FROM RP_入力帳票取得ヘッダ")
            .AppendLine($" WHERE 処理済={P}処理済")
            .GetListAsync<入力帳票取得ヘッダ>(new {
                処理済 = 入力帳票取得ヘッダ.処理未Val
            });

    /// <summary>
    /// RP_入力帳票取得明細テーブルのレコード取得
    /// </summary>
    /// <param name="id"></param>
    public async Task<IList<入力帳票取得明細>> Get入力帳票取得明細ListAsync(int? id) {
        if (id != null) {
            return await new StringBuilder()
                .AppendLine($"SELECT * FROM RP_入力帳票取得明細")
                .AppendLine($" WHERE ID={P}ID")
                .GetListAsync<入力帳票取得明細>(new {
                    ID = id
                });
        } else {
            return new List<入力帳票取得明細>();
        }
    }

    public static IList<string>? 作業完了製品ロットlist;
    public static string 作業完了製品ロット = "";

    #region 帳票データヘッダ
    /// <summary>
    /// RP_帳票データヘッダを更新する
    /// </summary>
    public async Task Update帳票データヘッダAsync(IList<int> idlist) {
        using var t = await _transactionFactory.CreateInstanceAsync();

        foreach (var id in idlist) {
            var sql = new StringBuilder()
                .AppendLine($"UPDATE RP_帳票データヘッダ")
                .AppendLine($"   SET 処理済={P}処理済")
                .AppendLine($"     , 処理日時=SYSDATETIME()")
                .AppendLine($" WHERE ID={P}ID")
                .ToString();

            await t.ExecuteAsyncEx(sql, param: new {
                処理済 = 帳票データヘッダ.処理済Val,
                ID = id
            });
        }
        t.Commit();
    }

    /// <summary>
    /// RP_帳票データヘッダのレコード取得
    /// </summary>
    public async Task<IList<帳票データヘッダ>> Get帳票データヘッダListAsync() {
        using var cn = await DbAccess.OpenAsync();

        var sql = new StringBuilder()
            .AppendLine($"SELECT * FROM RP_帳票データヘッダ")
            .AppendLine($" WHERE 処理済={P}処理済")
            .ToString();

        var list = await cn.QueryAsyncEx<帳票データヘッダ>(sql, new {
            処理済 = 帳票データヘッダ.処理未Val
        });

        if (list != null) {
            return list.ToList();
        } else {
            return new List<帳票データヘッダ>();
        }
    }

    /// <summary>
    /// 帳票データヘッダテーブルを削除する
    /// </summary>
    public async Task Delete帳票データヘッダAsync() {
        using var t = await _transactionFactory.CreateInstanceAsync();
        await Delete帳票データヘッダAsync(t);
        t.Commit();
    }
    public async Task Delete帳票データヘッダAsync(IMyTransaction t) {
        var sql = new StringBuilder()
            .AppendLine($"TRUNCATE TABLE RP_帳票データヘッダ")
            .AppendLine($";")
            .ToString();
        await t.ExecuteAsyncEx(sql);
    }
    #endregion
}
