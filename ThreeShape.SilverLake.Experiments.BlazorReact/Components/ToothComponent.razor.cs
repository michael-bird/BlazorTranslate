using Microsoft.AspNetCore.Components;

namespace ThreeShape.SilverLake.Experiments.BlazorReact.Components
{
    public partial class ToothComponent
    {
        [Parameter]
        public int ToothId { get; set; }

        [Parameter]
        public RenderFragment? ChildContent { get; set; }

        [Parameter]
        public EventCallback<Tuple<int, bool>> OnClick { get; set; }

        private bool _isCrown = false;

        private string _class { get { return _isCrown ? "crown" : ""; } }

        private void ToothClicked()
        {
            _isCrown = !_isCrown;
            OnClick.InvokeAsync(new Tuple<int, bool>(ToothId, _isCrown));
        }
    }
}
