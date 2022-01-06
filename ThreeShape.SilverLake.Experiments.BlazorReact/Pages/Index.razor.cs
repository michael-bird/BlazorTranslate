using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using MudBlazor;
using ThreeShape.SilverLake.Experiments.BlazorReact.Components;

namespace ThreeShape.SilverLake.Experiments.BlazorReact.Pages
{
    public partial class Index
    {
        [Inject]
        private IJSRuntime JsRuntime { get; set; }

        [Inject]
        private IDialogService DialogService { get; set; }

        private ElementReference _react;

        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            if (firstRender) await JsRuntime.InvokeVoidAsync("renderReact", _react);
        }

        private void OpenDialog()
        {
            var options = new MudBlazor.DialogOptions
            {
                CloseOnEscapeKey = true,
                FullWidth = true,
                CloseButton = true,
                MaxWidth = MaxWidth.Large
            };
            DialogService.Show<CaseEntryDialogComponent>("Case Entry", options);
        }
    }
}