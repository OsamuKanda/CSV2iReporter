using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Dapper;
using Dapper.Contrib;
using Microsoft.Extensions.Configuration;
using System.Data.Common;
using Npgsql;

namespace DbAccessLib; 
public static class MyDbFactory {
    public static bool IsMultiProvider { get; set; } = true;

    public static string DefaultProviderName { get; set; } = "Microsoft.Data.SqlClient";
    private static DbProviderFactory? singleFactory = null;

    public static DbProviderFactory GetFactory(string? providerInvariantName = null) {
        if (!IsMultiProvider) {
            if (singleFactory == null) {
                singleFactory = System.Data.SqlClient.SqlClientFactory.Instance;
            }
            return singleFactory;
        }
    
        if (string.IsNullOrEmpty(providerInvariantName)) {
            providerInvariantName = DefaultProviderName;
            //デフォルトのFactoryがなければ登録する。
            if (!DbProviderFactories.TryGetFactory(providerInvariantName, out var factory)) {
                DbProviderFactories.RegisterFactory(providerInvariantName, System.Data.SqlClient.SqlClientFactory.Instance);
            }
        } else {
            DbProviderFactories.RegisterFactory(providerInvariantName, Npgsql.NpgsqlFactory.Instance);
        }

        // register SqlClientFactory in provider factories
        return DbProviderFactories.GetFactory(providerInvariantName);
    }
}
public static class DbAccess {

    private static IConfigurationRoot? _configuration = null;

    public static async Task<IDbConnection> OpenAsync(string? conName = null, string? providerInvariantName = null) {
        var file = $"appsettings.json";
        var basepath = AppDomain.CurrentDomain.BaseDirectory;
        var filepath = $@"{basepath}";
        if (!File.Exists($@"{filepath}\{file}")) {
            file = "appsettings.json";
        }
        if (_configuration == null) {
            _configuration = new ConfigurationBuilder()
              .SetBasePath($@"{filepath}")
              .AddJsonFile(file, true, true)
              .Build();
        }

        var sec = _configuration
            .GetSection("Database");
        //var conName = sec["ConnectionName"];
        conName ??= "SqlServer";
        var conStr = sec
            .GetSection("ConnectionString")[conName];

        var factory = MyDbFactory.GetFactory(providerInvariantName);
        var connection = factory.CreateConnection();
        if (connection == null) {
            throw new Exception("factory.CreateConnection()に失敗しました。");
        }
        connection.ConnectionString = conStr;
        await connection.OpenAsync();

        return connection;
    }

    public static async Task<T> QueryFirstOrDefaultAsyncEx<T>(this IDbConnection cnn, string sql, object? param = null, IDbTransaction? transaction = null, int? commandTimeout = null, CommandType? commandType = null) {
        try {
            return await cnn.QueryFirstOrDefaultAsync<T>(sql, param, transaction, commandTimeout, commandType);
        } catch (Exception ex) {
            Serilog.Log.Logger.Error(ex, "例外発生");
            transaction?.Rollback();
            throw;
        }
    }

    public static async Task<IEnumerable<T>> QueryAsyncEx<T>(this IDbConnection cnn, string sql, object? param = null, IDbTransaction? transaction = null, int? commandTimeout = null, CommandType? commandType = null) {
        try {
            var result = await cnn.QueryAsync<T>(sql, param, transaction, commandTimeout, commandType);
            return result;
        } catch (Exception ex) {
            Serilog.Log.Logger.Error(ex, "例外発生");
            transaction?.Rollback();
            throw;
        }
    }

    public static async Task<int> ExecuteAsyncEx(this IDbConnection cnn, string sql, object? param = null, IDbTransaction? transaction = null, int? commandTimeout = null, CommandType? commandType = null) {
        try {
            var result = await cnn.ExecuteAsync(sql, param, transaction, commandTimeout, commandType);
            return result;
        } catch (Exception ex) {
            Serilog.Log.Logger.Error(ex, "例外発生");
            transaction?.Rollback();
            throw;
        }
    }

    public static async Task<IList<T>> GetListAsync<T>(this string sql, object? param = null, int? commandTimeout = null, string? conName = null, string? providerInvariantName = null) {
        using var cn = await DbAccess.OpenAsync(conName:conName, providerInvariantName:providerInvariantName);
        var q = param == null ?
            await cn.QueryAsyncEx<T>(sql, commandTimeout: commandTimeout) :
            await cn.QueryAsyncEx<T>(sql, param, commandTimeout: commandTimeout);
        if (q == null) {
            return new List<T>();
        }
        return q.ToList();
    }
    public static async Task<IList<T>> GetListAsync<T>(this StringBuilder sql, object? param = null, int? commandTimeout = null, string? conName = null, string? providerInvariantName = null) =>
        await GetListAsync<T>(sql.ToString(), param, commandTimeout:commandTimeout, conName:conName, providerInvariantName:providerInvariantName);

    public static async Task<T> QueryFirstOrDefaultAsyncEx<T>(this StringBuilder sql, object? param = null, IDbTransaction? transaction = null, int? commandTimeout = null, CommandType? commandType = null) {
        try {
            using var cn = await DbAccess.OpenAsync();
            return await cn.QueryFirstOrDefaultAsync<T>(sql.ToString(), param, transaction, commandTimeout, commandType);
        } catch (Exception ex) {
            Serilog.Log.Logger.Error(ex, "例外発生");
            transaction?.Rollback();
            throw;
        }
    }
}
