using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;

namespace GenerateDeliveryReports.Identity;

public static class DbInitializer
{
    /// <summary>
    /// Safely initializes the database, handling both fresh installs and schema updates on existing databases.
    /// Unlike Migrate() which fails if tables already exist, this catches "already exists" errors and treats them as success.
    /// This is needed because the project didn't start with migrations, so existing deployments have tables
    /// but no migration history, and new migrations assume a clean slate.
    /// </summary>
    public static void InitializeSafely(ApplicationDbContext db, ILogger logger)
    {
        try
        {
            db.Database.Migrate();
            logger.LogInformation("Database migration applied successfully.");
        }
        catch (Exception ex) when (ex.Message.Contains("already exists") || ex.Message.Contains("duplicate column"))
        {
            // On existing databases, new migrations fail because tables/columns already exist.
            // This is expected -- the schema is already there, just the migration history isn't tracked yet.
            logger.LogWarning("Database already has the schema; skipping migration. This is normal for existing deployments. Error: {Message}", ex.Message);
        }
    }
}
