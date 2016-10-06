using System.Collections.Generic;
using System.Threading.Tasks;

namespace SchneiderElectric.UnityComments
{
    public interface ITranslator
    {
        Task<string> Translate(string text, string from = "en", string to = "fr");
        Task<string> DetectSourceLanguage(string text);
        List<string> LanguagesCodes  { get; }
        List<string> LanguagesNames { get; }
        string LanguageNameFromCode(string name);
        string LanguageCodeFromName(string code);

    }
}