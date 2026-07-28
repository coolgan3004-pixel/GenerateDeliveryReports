using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;

namespace GenerateDeliveryReports.Identity;

/// <summary>
/// Satisfied only for accounts that have at least one linked external login (currently, Entra
/// via Microsoft.Identity.Web). Accounts created on the "create team login" admin page are local
/// password-only accounts with no external login record, so they can never satisfy this --
/// only whoever can complete the actual Entra SSO flow can reach pages that require it.
/// </summary>
public class ExternalLoginRequirement : IAuthorizationRequirement
{
}

public class ExternalLoginAuthorizationHandler : AuthorizationHandler<ExternalLoginRequirement>
{
    private readonly UserManager<ApplicationUser> _userManager;

    public ExternalLoginAuthorizationHandler(UserManager<ApplicationUser> userManager)
    {
        _userManager = userManager;
    }

    protected override async Task HandleRequirementAsync(AuthorizationHandlerContext context, ExternalLoginRequirement requirement)
    {
        if (context.User.Identity?.IsAuthenticated != true)
            return;

        var user = await _userManager.GetUserAsync(context.User);
        if (user == null)
            return;

        var logins = await _userManager.GetLoginsAsync(user);
        if (logins.Count > 0)
            context.Succeed(requirement);
    }
}
