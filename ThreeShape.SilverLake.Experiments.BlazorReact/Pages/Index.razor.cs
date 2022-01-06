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

        private DotNetObjectReference<Index>? objRef;

        private string _selectedItem;

        protected override async Task OnInitializedAsync()
        {
            _selectedItem = "Select an item using component and it will send selection back to dotNet and render it here.";
        }

        protected override async Task OnAfterRenderAsync(bool firstRender)
        {
            //Pass down to react component so it can call dotNet JSInvokable functions in this component
            objRef = DotNetObjectReference.Create(this);

            var items = new List<ToothProduct>
            {
                new ToothProduct{ Id= 1, Name = "Crown"},
                new ToothProduct{ Id= 2, Name = "Bridge"},
                new ToothProduct{ Id= 3, Name = "Additional"},
                new ToothProduct{ Id= 4, Name = "Implant"},
                new ToothProduct{ Id= 5, Name = "Coping"},
                new ToothProduct{ Id= 6, Name = "Multi-Crown"},
            };

            var serialisedItems = JsonConvert.SerializeObject(items);

            if (firstRender) await JsRuntime.InvokeVoidAsync("renderReact", objRef, _react, serialisedItems);
        }

        [JSInvokable("SetSelected")]
        public void SetSelectedItem( string item)
        {
            _selectedItem = item;
            StateHasChanged();
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