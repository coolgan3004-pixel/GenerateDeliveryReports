namespace GenerateDeliveryReports.Data.Services
{
    public class BriefingCache
    {
        public record BriefingSnapshot(DateTime GeneratedAt, string Html, string Markdown);

        private readonly object _lock= new();

        private BriefingSnapshot? _latest;

        public BriefingSnapshot? Latest { get 
        {
            lock (_lock)
            return _latest;
            
        }}

        public void Set(BriefingSnapshot snapshot)
        {
            lock (_lock)
            _latest = snapshot;
        }
    }
}
