using Newtonsoft.Json;

namespace ThreeShape.SilverLake.Experiments.BlazorReact.Models
{
    public class Doctor
    {
        public int Id { get; set; }

        public int ClientId { get; set; }

        public string Name { get; set; }

        [JsonProperty("first_name")]
        public string FirstName { get; set; }

        [JsonProperty("last_name")]
        public string LastName { get; set; }
    }
}
