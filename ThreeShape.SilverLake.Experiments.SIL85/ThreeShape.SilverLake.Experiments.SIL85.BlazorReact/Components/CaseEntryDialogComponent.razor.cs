using Microsoft.AspNetCore.Components;
using MudBlazor;

namespace ThreeShape.SilverLake.Experiments.SIL85.BlazorReact.Components
{
    public partial class CaseEntryDialogComponent
    {
        [CascadingParameter] 
        MudDialogInstance MudDialog { get; set; }

        void Cancel() => MudDialog.Cancel();
    }
}
