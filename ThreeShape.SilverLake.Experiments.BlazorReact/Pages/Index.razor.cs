using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using MudBlazor;
using Newtonsoft.Json;
using ThreeShape.SilverLake.Experiments.BlazorReact.Components;
using ThreeShape.SilverLake.Experiments.BlazorReact.Models;

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
            var items = new List<ToothProduct>
            {
                new ToothProduct{ Id= 1, Name = "Crown"},
                new ToothProduct{ Id= 2, Name = "Bridge"},
                new ToothProduct{ Id= 3, Name = "Additional"},
                new ToothProduct{ Id= 4, Name = "Implant"},
                new ToothProduct{ Id= 5, Name = "Coping"},
                new ToothProduct{ Id= 5, Name = "Multi-Crown"},
            };

            var serialisedItems = JsonConvert.SerializeObject(items);

            if (firstRender) await JsRuntime.InvokeVoidAsync("renderReact", _react, serialisedItems);
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