using AutoReportLib.Interfaces;
using DbAccessLib;
using System.Data;

namespace AutoReportLib.Services;
public class MyTransaction : IMyTransaction {
    private IDbConnection Connection { get; set; } = null!;
    private IDbTransaction Transaction { get; set; } = null!;

    private MyTransaction() { }

    public async static Task<IMyTransaction> CreateAndOpenAndBeginTransaction() {
        var o = new MyTransaction();
        await o.OpenAndBeginTransaction();
        return o;
    }

    private async Task OpenAndBeginTransaction() {
        Connection = await DbAccess.OpenAsync();
        Transaction = Connection.BeginTransaction();
    }

    public void Commit() =>
        Transaction.Commit();

    public void Dispose() {
        Transaction.Dispose();
        Connection.Dispose();
    }

    public async Task<int> ExecuteAsyncEx(string sql, object? param = null, int? commandTimeout = null, CommandType? commandType = null) =>
        await Connection.ExecuteAsyncEx(sql, param, transaction: Transaction, commandTimeout, commandType);

    public async Task<IEnumerable<T>> QueryAsyncEx<T>(string sql, object? param = null, int? commandTimeout = null, CommandType? commandType = null) =>
        await Connection.QueryAsyncEx<T>(sql, param, Transaction, commandTimeout, commandType);

    public async Task<T> QueryFirstOrDefaultAsyncEx<T>(string sql, object? param = null, int? commandTimeout = null, CommandType? commandType = null) =>
        await Connection.QueryFirstOrDefaultAsyncEx<T>(sql, param, Transaction, commandTimeout, commandType);
}

public class MyTransactionFactory : IMyTransactionFactory {
    public async Task<IMyTransaction> CreateInstanceAsync() =>
         await MyTransaction.CreateAndOpenAndBeginTransaction();
}
