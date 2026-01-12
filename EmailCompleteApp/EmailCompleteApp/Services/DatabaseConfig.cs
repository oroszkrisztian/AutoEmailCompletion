using System;
using Microsoft.EntityFrameworkCore;
using Npgsql; // from Npgsql package
using EmailCompleteApp.Data;

namespace EmailCompleteApp.Services
{
   
    public static class DatabaseConfig
    {
        

        private const string DatabaseUri = "postgresql://postgres.mvgfbmdgnzfisuacduqw:41VR1fwZXvtqpAcr@aws-1-eu-north-1.pooler.supabase.com:5432/postgres";

        public static string RawUri => DatabaseUri;

        public static string ConnectionString => BuildConnectionString(RawUri);

        public static DbContextOptions<AppDbContext> GetDbContextOptions()
        {
            var optionsBuilder = new DbContextOptionsBuilder<AppDbContext>();

            optionsBuilder.UseNpgsql(ConnectionString, npgsqlOptions =>
            {
                npgsqlOptions.EnableRetryOnFailure(
                    maxRetryCount: 5,
                    maxRetryDelay: TimeSpan.FromSeconds(30),
                    errorCodesToAdd: null);

                npgsqlOptions.CommandTimeout(60);
            });

#if DEBUG
            optionsBuilder.EnableSensitiveDataLogging();
            optionsBuilder.EnableDetailedErrors();
            optionsBuilder.LogTo(message => System.Diagnostics.Debug.WriteLine(message));
#endif

            return optionsBuilder.Options;
        }

        public static AppDbContext CreateDbContext() => new(GetDbContextOptions());

        private static string BuildConnectionString(string uriOrConn)
        {
            if (string.IsNullOrWhiteSpace(uriOrConn))
                throw new InvalidOperationException("No database URI provided. Set SUPABASE_DATABASE_URL or update DatabaseConfig.FallbackUri.");

            if (uriOrConn.Contains("=") && uriOrConn.Contains(";")) return uriOrConn;

            if (!(uriOrConn.StartsWith("postgres://", StringComparison.OrdinalIgnoreCase) ||
                  uriOrConn.StartsWith("postgresql://", StringComparison.OrdinalIgnoreCase)))
            {
                return uriOrConn;
            }

            var uri = new Uri(uriOrConn);
            var userInfo = uri.UserInfo.Split(':', 2);
            var username = userInfo.Length > 0 ? Uri.UnescapeDataString(userInfo[0]) : "postgres";
            var password = userInfo.Length > 1 ? Uri.UnescapeDataString(userInfo[1]) : string.Empty;
            var host = uri.Host;
            var port = uri.Port > 0 ? uri.Port : 5432;
            var database = uri.AbsolutePath?.TrimStart('/') ?? "postgres";

            var builder = new NpgsqlConnectionStringBuilder
            {
                Host = host,
                Port = port,
                Username = username,
                Password = password,
                Database = database,
                SslMode = SslMode.Require,
                TrustServerCertificate = true,
            };

            return builder.ToString();
        }

       
    }
}