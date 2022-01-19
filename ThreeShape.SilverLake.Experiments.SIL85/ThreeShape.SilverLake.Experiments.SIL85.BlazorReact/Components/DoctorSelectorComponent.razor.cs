using Microsoft.AspNetCore.Components;
using ThreeShape.SilverLake.Experiments.SIL85.BlazorReact.Models;
using ThreeShape.SilverLake.Experiments.SIL85.BlazorReact.Services;

namespace ThreeShape.SilverLake.Experiments.SIL85.BlazorReact.Components
{
    public partial class DoctorSelectorComponent
    {
        [Inject] LabstarService _labstarService { get; set; }

        private Doctor selectedDoctor;

        private IEnumerable<Doctor> Doctors;

        protected override async Task OnInitializedAsync()
        {
            try
            {
                Doctors = await _labstarService.GetDoctorsAsync();
            }
            catch (Exception ex)
            {
                //For debugging this writes to browser console in WebAssembly
                Console.WriteLine(ex);
            }
        }

        private async Task<IEnumerable<Doctor>> Search(string value)
        {
            if (string.IsNullOrEmpty(value) && Doctors.Any())
                return Doctors;
            if (value.Equals(selectedDoctor.Name) && Doctors.Any())
                return Doctors;
            return Doctors.Where(x => x.Name.Contains(value, StringComparison.InvariantCultureIgnoreCase)) ?? Enumerable.Empty<Doctor>();

        }

    }
}
