using Microsoft.AspNetCore.Components;

namespace ThreeShape.SilverLake.Experiments.SIL85.BlazorReact.Components
{
    public partial class ToothArchComponent
    {
        [Parameter]
        public string Class { get; set; }

        [Parameter]
        public RenderFragment? ChildContent { get; set; }

        private string _baseClass = "arch_wrapper";

        private string _class { get { return $"{Class} {_baseClass}"; } }
    }
}
