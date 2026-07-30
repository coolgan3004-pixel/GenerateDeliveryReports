using System.Security.Claims;
using GenerateDeliveryReports.Identity;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace GenerateDeliveryReports.Areas.Identity.Pages.Account;

/// <summary>
/// Overrides Identity UI's compiled ExternalLogin page (same route: /Identity/Account/ExternalLogin
/// -- an app page at the same path takes precedence over the Razor Class Library's). The stock
/// page always stops on a "Register" form asking the user to type their email before it will
/// create the local account and finish signing in, even though Entra already gave us one via
/// claims. This auto-provisions (or links) the local ApplicationUser from that claim instead, so a
/// first-time Entra sign-in lands the user straight in the app -- the manual form is kept only as
/// a fallback for the rare case where no usable email-shaped claim comes back from Entra.
///
/// [AllowAnonymous] is required explicitly here: Identity's own convention that exempts the whole
/// /Account folder from the app's global FallbackPolicy (RequireAuthenticatedUser) is applied to
/// the Razor Class Library's compiled pages and did not reliably extend to this app-level page
/// overriding one of them once published -- confirmed live (Azure App Service, published build):
/// an anonymous POST here was being redirected straight to /Identity/Account/Login by the
/// authorization middleware before OnPost (the Entra challenge) ever ran, producing an infinite
/// Login&lt;-&gt;ExternalLogin bounce with the returnUrl nesting deeper on every click.
/// </summary>
[AllowAnonymous]
public class ExternalLoginModel : PageModel
{
    private readonly SignInManager<ApplicationUser> _signInManager;
    private readonly UserManager<ApplicationUser> _userManager;
    private readonly ILogger<ExternalLoginModel> _logger;

    public ExternalLoginModel(SignInManager<ApplicationUser> signInManager, UserManager<ApplicationUser> userManager, ILogger<ExternalLoginModel> logger)
    {
        _signInManager = signInManager;
        _userManager = userManager;
        _logger = logger;
    }

    [BindProperty]
    public InputModel Input { get; set; } = new();

    public string ProviderDisplayName { get; set; } = string.Empty;

    public string ReturnUrl { get; set; } = string.Empty;

    [TempData]
    public string? ErrorMessage { get; set; }

    public class InputModel
    {
        [System.ComponentModel.DataAnnotations.Required]
        [System.ComponentModel.DataAnnotations.EmailAddress]
        public string Email { get; set; } = string.Empty;
    }

    public IActionResult OnGet() => RedirectToPage("./Login");

    public IActionResult OnPost(string provider, string? returnUrl = null)
    {
        var redirectUrl = Url.Page("./ExternalLogin", pageHandler: "Callback", values: new { returnUrl });
        var properties = _signInManager.ConfigureExternalAuthenticationProperties(provider, redirectUrl);
        return new ChallengeResult(provider, properties);
    }

    public async Task<IActionResult> OnGetCallbackAsync(string? returnUrl = null, string? remoteError = null)
    {
        returnUrl ??= Url.Content("~/");
        if (remoteError != null)
        {
            ErrorMessage = $"Error from external provider: {remoteError}";
            return RedirectToPage("./Login", new { ReturnUrl = returnUrl });
        }

        var info = await _signInManager.GetExternalLoginInfoAsync();
        if (info == null)
        {
            ErrorMessage = "Error loading external login information.";
            return RedirectToPage("./Login", new { ReturnUrl = returnUrl });
        }

        // Already linked -- sign in directly. bypassTwoFactor: true because a second local factor
        // is redundant on top of whatever MFA Entra itself already enforced for this sign-in.
        var signInResult = await _signInManager.ExternalLoginSignInAsync(info.LoginProvider, info.ProviderKey, isPersistent: false, bypassTwoFactor: true);
        if (signInResult.Succeeded)
        {
            _logger.LogInformation("{Name} logged in with {LoginProvider} provider.", info.Principal.Identity?.Name, info.LoginProvider);
            // Full page load (Redirect, not LocalRedirect) to restart the Blazor circuit and refresh
            // its Task<AuthenticationState>. See TryProvisionAndSignInAsync for detailed explanation.
            return Redirect(returnUrl);
        }
        if (signInResult.IsLockedOut)
        {
            return RedirectToPage("./Lockout");
        }

        // First time this external login has been seen -- auto-provision (or link to an existing
        // matching local account) from the claims Entra already gave us, instead of making the
        // user type their email in manually.
        var email = info.Principal.FindFirstValue(ClaimTypes.Email)
            ?? info.Principal.FindFirstValue("preferred_username")
            ?? info.Principal.FindFirstValue(ClaimTypes.Upn)
            ?? info.Principal.FindFirstValue(ClaimTypes.Name);

        if (!string.IsNullOrWhiteSpace(email) && email.Contains('@') && await TryProvisionAndSignInAsync(email, info, returnUrl) is { } provisioned)
        {
            return provisioned;
        }

        // Fallback: no usable claim came back -- ask for the email manually, same as stock Identity.
        ProviderDisplayName = info.ProviderDisplayName ?? info.LoginProvider;
        ReturnUrl = returnUrl;
        if (info.Principal.FindFirstValue(ClaimTypes.Email) is { } claimEmail)
            Input = new InputModel { Email = claimEmail };
        return Page();
    }

    public async Task<IActionResult> OnPostConfirmationAsync(string? returnUrl = null)
    {
        returnUrl ??= Url.Content("~/");
        var info = await _signInManager.GetExternalLoginInfoAsync();
        if (info == null)
        {
            ErrorMessage = "Error loading external login information during confirmation.";
            return RedirectToPage("./Login", new { ReturnUrl = returnUrl });
        }

        if (!ModelState.IsValid)
        {
            ProviderDisplayName = info.ProviderDisplayName ?? info.LoginProvider;
            ReturnUrl = returnUrl;
            return Page();
        }

        var result = await TryProvisionAndSignInAsync(Input.Email, info, returnUrl);
        if (result != null)
            return result;

        ModelState.AddModelError(string.Empty, "Error creating the account for this email address.");
        ProviderDisplayName = info.ProviderDisplayName ?? info.LoginProvider;
        ReturnUrl = returnUrl;
        return Page();
    }

    private async Task<IActionResult?> TryProvisionAndSignInAsync(string email, ExternalLoginInfo info, string returnUrl)
    {
        var user = await _userManager.FindByEmailAsync(email);
        if (user == null)
        {
            user = new ApplicationUser { UserName = email, Email = email, EmailConfirmed = true };
            var createResult = await _userManager.CreateAsync(user);
            if (!createResult.Succeeded)
            {
                _logger.LogWarning("Failed to auto-provision account for {Email} via {LoginProvider}: {Errors}",
                    email, info.LoginProvider, string.Join(" ", createResult.Errors.Select(e => e.Description)));
                return null;
            }
        }

        var addLoginResult = await _userManager.AddLoginAsync(user, info);
        if (!addLoginResult.Succeeded)
        {
            _logger.LogWarning("Failed to link {LoginProvider} login to {Email}: {Errors}",
                info.LoginProvider, email, string.Join(" ", addLoginResult.Errors.Select(e => e.Description)));
            return null;
        }

        await _signInManager.SignInAsync(user, isPersistent: false, info.LoginProvider);
        _logger.LogInformation("Auto-provisioned/linked {Email} via {LoginProvider}.", email, info.LoginProvider);
        // Redirect must be a full page load (not LocalRedirect) to restart the Blazor circuit and
        // refresh its Task<AuthenticationState>. LocalRedirect stays on the same circuit, which was
        // initialized before the sign-in and still sees the user as unauthenticated -- the cookies
        // are set correctly, but <AuthorizeView> components (like NavMenu's logout button) won't see
        // the change until the circuit restarts.
        return Redirect(returnUrl);
    }
}
