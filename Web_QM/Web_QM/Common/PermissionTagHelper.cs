using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc.TagHelpers;
using Microsoft.AspNetCore.Razor.TagHelpers;

namespace Web_QM.Common
{
    [HtmlTargetElement("a", Attributes = "asp-policy")]
    public class PermissionTagHelper : TagHelper
    {
        private readonly IAuthorizationService _authorizationService;
        private readonly IHttpContextAccessor _httpContextAccessor;

        public PermissionTagHelper(IAuthorizationService authorizationService, IHttpContextAccessor httpContextAccessor)
        {
            _authorizationService = authorizationService;
            _httpContextAccessor = httpContextAccessor;
        }

        [HtmlAttributeName("asp-policy")]
        public string PolicyName { get; set; }

        public override async Task ProcessAsync(TagHelperContext context, TagHelperOutput output)
        {
            var httpContext = _httpContextAccessor.HttpContext;
            var result = await _authorizationService.AuthorizeAsync(httpContext.User, PolicyName);

            if (!result.Succeeded)
            {
                output.SuppressOutput();
            }
        }
    }
}

