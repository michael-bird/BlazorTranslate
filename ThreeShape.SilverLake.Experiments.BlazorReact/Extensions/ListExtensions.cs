using System.Text;

namespace ThreeShape.SilverLake.Experiments.BlazorReact.Extensions
{
    public static class ListExtensions
    {
        public static string ToFormattedString(this IList<int> collection)
        {
            int range = 0;
            var ordered = collection.OrderBy(x => x).ToList();
            StringBuilder sb = new StringBuilder();

            if (ordered.Count > 0)
            {
                sb.Append(ordered[0]);

                for (int i = 1; i < ordered.Count; i++)
                {
                    if (ordered[i] - ordered[i - 1] == 1)
                    {
                        range = ordered[i];
                    }
                    else
                    {
                        if (range > 0)
                        {
                            sb.Append($"-{range}");
                            range = 0;
                        }

                        sb.Append($",{ordered[i]}");
                    }
                }
                if (range > 0)
                {
                    sb.Append($"-{range}");
                }
            }
            else
            {
                sb.Append('#');
            }

            return sb.ToString();
        }
    }
}
