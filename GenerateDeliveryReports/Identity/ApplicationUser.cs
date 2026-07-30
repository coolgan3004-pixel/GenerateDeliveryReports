using Microsoft.AspNetCore.Identity;

namespace GenerateDeliveryReports.Identity;

public class ApplicationUser : IdentityUser
{
    public bool IsActive { get; set; } = true;
    public DateTime CreatedUtc { get; set; }
    public DateTime? UpdatedUtc { get; set; }
    public DateTime? DeactivatedUtc { get; set; }
}
