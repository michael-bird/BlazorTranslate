using Newtonsoft.Json.Linq;
using ThreeShape.SilverLake.Experiments.BlazorReact.Models;

namespace ThreeShape.SilverLake.Experiments.BlazorReact.Services
{
    public class LabstarService
    {
        private readonly HttpClient _httpClient;

        public LabstarService(HttpClient httpClient)
        {
            _httpClient = httpClient;

            _httpClient.BaseAddress = new Uri("http://localhost:3000");
        }

        public async Task<IEnumerable<Doctor>> GetDoctorsAsync()
        {
            var response = await _httpClient.GetStringAsync("/doctor?extraFields=notation,clientInfo,orderingPrompt&searchField=clients.bActive&searchValue=1&limit=-1");

            JObject responseObject = JObject.Parse(response);

            var results = responseObject["data"].Children().ToList();

            return results.Select((doctor) => doctor.ToObject<Doctor>());
        }

    }
}
