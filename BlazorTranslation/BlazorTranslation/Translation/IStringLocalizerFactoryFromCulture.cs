using Microsoft.Extensions.Localization;

namespace BlazorTranslation.Translation
{
    public interface IStringLocalizerFactoryFromCulture : IStringLocalizerFactory
    {
        IStringLocalizer Create();
    }
}
