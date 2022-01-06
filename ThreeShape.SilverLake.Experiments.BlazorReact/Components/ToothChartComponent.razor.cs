using ThreeShape.SilverLake.Experiments.BlazorReact.Extensions;

namespace ThreeShape.SilverLake.Experiments.BlazorReact.Components
{
    public partial class ToothChartComponent
    {
        private List<int> SelectedTeeth;
        private string _toothSelection = "#";

        protected override void OnInitialized()
        {
            SelectedTeeth = new List<int>();
        }

        private void ClickHandler(Tuple<int, bool> toothInfo)
        {
            if (toothInfo.Item2)
            {
                if (!SelectedTeeth.Contains(toothInfo.Item1))
                {
                    SelectedTeeth.Add(toothInfo.Item1);
                }
            }
            else
            {
                if (SelectedTeeth.Contains(toothInfo.Item1))
                {
                    SelectedTeeth.Remove(toothInfo.Item1);
                }
            }

            ShowSelectedTeeth();
        }

        private void ShowSelectedTeeth()
        {
            _toothSelection = SelectedTeeth.ToFormattedString();
        }
    }
}
