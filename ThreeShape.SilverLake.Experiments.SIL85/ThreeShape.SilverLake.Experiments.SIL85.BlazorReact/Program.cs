using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using MudBlazor.Services;
using ThreeShape.SilverLake.Experiments.SIL85.BlazorReact;
using ThreeShape.SilverLake.Experiments.SIL85.BlazorReact.Services;

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

builder.Services.AddHttpClient<LabstarService>();
builder.Services.AddMudServices();

await builder.Build().RunAsync();
