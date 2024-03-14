//using System;
//using System.Collections.Generic;
using System.Data;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

namespace AutoReportLib.Interfaces;

public interface IMyTransactionFactory {
    Task<IMyTransaction> CreateInstanceAsync();
}

public interface IMyTransaction : IDisposable {
    void Commit();

    Task<int> ExecuteAsyncEx(string sql, object? param = null, int? commandTimeout = null, CommandType? commandType = null);
    Task<IEnumerable<T>> QueryAsyncEx<T>(string sql, object? param = null, int? commandTimeout = null, CommandType? commandType = null);
    Task<T> QueryFirstOrDefaultAsyncEx<T>(string sql, object? param = null, int? commandTimeout = null, CommandType? commandType = null);
}
