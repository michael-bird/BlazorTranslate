using BlazorTranslation.Translation;
using Microsoft.AspNetCore.Components;
using Microsoft.Extensions.Localization;

namespace BlazorTranslation.Pages
{
    public partial class Index
    {

        [Inject] IStringLocalizerFactoryFromCulture _localizerFactory { get; set; }

        private IStringLocalizer _localizer => _localizerFactory.Create();


        protected override async Task OnInitializedAsync()
        {

        }

    }
}
