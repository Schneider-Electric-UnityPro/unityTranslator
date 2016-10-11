using log4net;
using Microsoft.Crm.UnifiedServiceDesk.Dynamics.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace SchneiderElectric.UnityComments
{

    public abstract class Translator : ITranslator
    {

        #region fields
        protected static string _credential;
        protected static Dictionary<string, string> _languageToCode;
        protected static Dictionary<string, string> _codeToLanguage;
        #endregion

        #region Prperties
        /// <summary>
        /// Gets or sets the log.
        /// </summary>
        /// <value>
        /// The log.
        /// </value>
        public static ILog Log {get;set;}
        /// <summary>
        /// Gets the supported languages codes.
        /// </summary>
        /// <value>
        /// The languages codes.
        /// </value>
        virtual public List<string> LanguagesCodes => _languageToCode?.Values.ToList();
        /// <summary>
        /// Gets  the supported languages names.
        /// </summary>
        /// <value>
        /// The languages names.
        /// </value>
        virtual public List<string> LanguagesNames => _codeToLanguage?.Values.ToList();

        #endregion
        /// <summary>
        /// Name of Languages from code.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns></returns>
        public string LanguageNameFromCode(string name)
        {
            return _codeToLanguage[name];
        }

        /// <summary>
        /// code of Languages from name.
        /// </summary>
        /// <param name="code">The code.</param>
        /// <returns></returns>
        public string LanguageCodeFromName(string code)
        {
            return _languageToCode[code];
        }

        /// <summary>
        /// Translates the specified text.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <param name="from">From.</param>
        /// <param name="to">To.</param>
        /// <returns></returns>
        abstract public Task<string> Translate(string text, string from = "en", string to = "fr");
        /// <summary>
        /// Detects the source language.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <returns></returns>
        abstract public Task<string> DetectSourceLanguage(string text);
    }


    /// <summary>
    /// Text translation using microsoft.com web services
    /// current  choice
    /// quality of trad  : pretty good
    /// </summary>
    public class MicrosoftTranslator : Translator
    {

        class CloudIdentifier
        {
            private const string MicrosoftTranslatorAccountkey = "MicrosoftTranslatorAccountkey";
            private const string MicrosoftTranslatorAccountPwd= "MicrosoftTranslatorAccountPwd";
            /// <summary>
            /// Gets or sets the key.
            /// </summary>
            /// <value>
            /// The key.
            /// </value>
            public string Key
            {
                get; private set;
            }
            /// <summary>
            /// Gets or sets the password.
            /// </summary>
            /// <value>
            /// The password.
            /// </value>
            public string Password
            {
                get; private set;
            }

            /// <summary>
            /// Gets or sets a value indicating whether this <see cref="CloudIdentifier" /> is valid.
            /// </summary>
            /// <value>
            ///   <c>true</c> if valid; otherwise, <c>false</c>.
            /// </value>
            public bool Valid => !string.IsNullOrEmpty(Key) && !string.IsNullOrEmpty(Password);

            /// <summary>
            /// Initializes a new instance of the <see cref="CloudIdentifier" /> class.
            /// </summary>
            public   CloudIdentifier(int index)
            {

                Key = System.Configuration.ConfigurationManager.AppSettings[$"{MicrosoftTranslatorAccountkey}{index}"];
                Password = System.Configuration.ConfigurationManager.AppSettings[$"{MicrosoftTranslatorAccountPwd}{index}"];
            }
        }

        #region private
        private static int _index = 0;

        //manages till 4 differents accounts. to be set in unitycomments.dll.config
        private static List<CloudIdentifier> CloudIds = new List<CloudIdentifier>()
        {
            new CloudIdentifier(1),
            new CloudIdentifier(2),
            new CloudIdentifier(3),
            new CloudIdentifier(4)
        };

        /// <summary>
        /// Gets the credentials.
        /// </summary>
        /// <value>
        /// The credentials.
        /// </value>
        private static CloudIdentifier CloudID => CloudIds[_index];

        /// <summary>
        /// Finds the next credentials.
        /// </summary>
        /// <returns></returns>
        private static bool FindNextCredentials()
        {
            int cur = _index;
            for (int i = 1; i < CloudIds.Count; i++)
            {
                int rank = (_index + i) % CloudIds.Count;
                if (CloudIds[rank].Valid)
                {
                    _index = rank;
                    return true;
                }
            }
            return false;
        }



        private static string _baseUri = "https://api.microsofttranslator.com/v2/Http.svc";       
        #endregion

        /// <summary>
        /// Initializes a new instance of the <see cref="MicrosoftTranslator" /> class.
        /// </summary>
        public MicrosoftTranslator()
        {
        }

        /// <summary>
        /// Gets or sets a value indicating whether this instance is initialised.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this instance is initialised; otherwise, <c>false</c>.
        /// </value>
        public static bool IsInitialised => !string.IsNullOrEmpty(_credential);

        /// <summary>
        /// Gets the supported languages codes.
        /// </summary>
        /// <value>
        /// The languages codes.
        /// </value>
        override public List<string> LanguagesCodes
        {
            get
            {
                if (_languageToCode == null)
                {
                    Initlanguages().Wait();
                }
                return base.LanguagesCodes;
            }
        }
        /// <summary>
        /// Gets  the supported languages names.
        /// </summary>
        /// <value>
        /// The languages names.
        /// </value>
        override public List<string> LanguagesNames
        {
            get
            {
                if (_languageToCode == null)
                {
                    Initlanguages().Wait();
                }           
                return base.LanguagesNames;
            }
        }

        /// <summary>
        /// Detects the source language.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <returns></returns>
        public override  async Task<string> DetectSourceLanguage(string text)
        {
            try
            {
                return await HttpGet<string>(BuildDetectUri(text));
            }
            catch (WebException e)
            {
                Log?.Error(e.Message);
                return string.Empty;
            }
        }

        /// <summary>
        /// Translates the specified text.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <param name="from">From.</param>
        /// <param name="to">To.</param>
        /// <returns></returns>
        public override async Task<string> Translate(string text, string from = "en", string to = "fr")
        {
            string translation = text;
            if (text != null)
            {
                if (LanguagesCodes.Contains(from) && LanguagesCodes.Contains(to))
                {
                    List<string> parts = Split(text);
                    StringBuilder Responsebuilder = new StringBuilder();
                    foreach (string s in parts)
                    {
                        Responsebuilder.Append( await HttpGet<string>(BuildTranslateUri(text, from, to)));
                    }
                    translation = Responsebuilder.ToString();
                    if (translation != null)
                    {
                        Log?.Debug($"Translation for source text '{text}' from {from} to {to} is {translation}");
                    }
                    else
                    {
                        Log?.Warn($"Translation for source text '{text}' from {from} to {to} does not return a valid string");
                        translation = text;
                    }
                }
                else
                {
                    throw new ApplicationException($"Error: languages {to} or {from} not suported");
                }
            }
            return translation;
        }

        /// <summary>
        /// Splits the specified text.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <returns></returns>
        private static List<string> Split(string text)
        {
            var maxlength = 250;
            text = text.Replace('\r', ' ').Replace('\n',' ');
            List<string> parts = new List<string>();
            if (text.Length > maxlength)
            {
                char[] seps = { ',', '.', ';'};
                parts = text.Split(seps,10 ,StringSplitOptions.RemoveEmptyEntries).ToList();
            }
            else
            {
                //one shot 
                parts.Add(text);
            }

            return parts;
        }

        #region private
        /// <summary>
        /// Sends the specified URL.
        /// </summary>
        /// <param name="url">The URL.</param>
        /// <returns>T</returns>
        /// <exception cref="WebException">no credentials</exception>
        private static async Task<T> HttpGet<T>(string url)
        {
            var retry = CloudIds.Count;//max retry 
            while (retry>0)
            {
                try
                {
                    retry--;
                    Initialize();
                    HttpWebRequest request = CreateRequest(url);
                    return await GetResponse<T>(request);
                }
                catch(Exception e)
                {
                   if(! CheckCredentials(true))
                   {
                        throw new WebException("Invalid Credential for Microsoft Translator service", e);
                    }
                   _credential = null; //forces the recalculation of the token
                }
            }
            throw new WebException("No valid Credential for Microsoft Translator service foud");
        }


        /// <summary>
        /// HTTP post.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="uri">The URI.</param>
        /// <param name="param">The parameter.</param>
        /// <returns></returns>
        private static async Task<T> HttpPost<T>(string uri, object param = null)
        {
            var retry = CloudIds.Count;//max retry 
            while (retry > 0)
            {
                try
                {
                    retry--;
                    Initialize();
                    // create the request
                    HttpWebRequest request = CreateRequest(uri);
                    SetPostParameters(param, request);
                    return await GetResponse<T>(request);
                }
                catch (Exception e)
                {
                    if (!CheckCredentials(true))
                    {
                        throw new WebException("Invalid Credential for Microsoft Translator service", e);
                    }
                    _credential = null; //forces the recalculation of the token
                }
            }
            throw new WebException("No valid Credential for Microsoft Translator service foud");
        }

        /// <summary>
        /// Gets the response.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="request">The request.</param>
        /// <returns></returns>
        private static async Task<T> GetResponse<T>(HttpWebRequest request)
        {
            using (WebResponse response = await request.GetResponseAsync())
            {
                return InterpretResponse<T>(response);
            }
        }

        /// <summary>
        /// Sets the post parameters.
        /// </summary>
        /// <param name="param">The parameter.</param>
        /// <param name="request">The request.</param>
        private static void SetPostParameters(object param, HttpWebRequest request, string contentType= "text/xml")
        {
            request.ContentType = contentType;
            request.Method = "POST";
            DataContractSerializer postParam = new DataContractSerializer(param.GetType());
            using (Stream stream = request.GetRequestStream())
            {
                postParam.WriteObject(stream, param);
            }
        }


        /// <summary>
        /// Interprets the response.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="response">The response.</param>
        /// <returns></returns>
        private static T InterpretResponse<T>(WebResponse response)
        {
            using (Stream stream = response.GetResponseStream())
            {
                DataContractSerializer dcs = new DataContractSerializer(typeof(T));
                return (T)dcs.ReadObject(stream);
            }
        }

        /// <summary>
        /// Creates the request.
        /// </summary>
        /// <param name="url">The URL.</param>
        /// <returns></returns>
        private static HttpWebRequest CreateRequest(string url)
        {
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.Headers.Add("Authorization", _credential);
            return httpWebRequest;
        }

        /// <summary>
        /// Builds the URI.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <param name="from">From.</param>
        /// <param name="to">To.</param>
        /// <returns></returns>
        private static string BuildTranslateUri(string text, string from, string to)
        {
            return _baseUri + "/Translate" + $"?text={HttpUtility.UrlEncode(text)}" + "&from=" + from + "&to=" + to;
        }

        /// <summary>
        /// Builds the detect URI.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <returns></returns>
        private static string BuildDetectUri(string text)
        {
            return _baseUri + "/Detect" + $"?text={HttpUtility.UrlEncode(text)}";
        }

        /// <summary>
        /// Requests the token.
        /// </summary>
        public static void Initialize()
        {
            if (!IsInitialised)
            {
                AdmAccessToken admToken;
                //Get Client Id and Client Secret from https://datamarket.azure.com/developer/applications/
                //Refer obtaining AccessToken (http://msdn.microsoft.com/en-us/library/hh454950.aspx) 

                if (CheckCredentials())
                {
                    AdmAuthentication admAuth = new AdmAuthentication(CloudID.Key, CloudID.Password);
                    try
                    {
                        admToken = admAuth.GetAccessToken();
                        // Create a header with the access_token property of the returned token
                        _credential = "Bearer " + admToken.access_token;
                    }
                    catch (WebException we)
                    {
                        Log?.Error(we);
                    }
                    catch (Exception e)
                    {
                        Log?.Error(e.Message);
                    }
                }
                else
                {
                    Log.Error("no valid credential found");
                }

            }
        }

        /// <summary>
        /// Checks the credentials.
        /// </summary>
        /// <param name="forceIncr">if set to <c>true</c> [force incr].</param>
        /// <returns> true if valid credential found</returns>
        private static bool CheckCredentials(bool forceIncr = false)
        {
            if (CloudIds.Any(x => x.Valid))
            {
                if (forceIncr)
                {
                    _index = (_index + 1) % CloudIds.Count;
                }

                while (!CloudID.Valid)
                {
                    if (!FindNextCredentials())
                    {
                        LogCredentialProblem();
                        break;
                    }
                }
            }
            else
            {
                LogCredentialProblem();
            }
            return CloudID.Valid;
        }

        /// <summary>
        /// Logs the credential problem.
        /// </summary>
        private static void LogCredentialProblem()
        {
            Log?.Error($"No credential founds for Microsoft translator - please update {AppDomain.CurrentDomain.SetupInformation.ConfigurationFile} with valid credentials");
            Log?.Info("Get Client Id and Client Secret from https://datamarket.azure.com/developer/applications");
            Log?.Info("Refer obtaining AccessToken (http://msdn.microsoft.com/en-us/library/hh454950.aspx)");
            Log?.Info("this tool support till x=1..4 microsoft accounts");
            Log?.Info($"put the client ID in {Path.GetFileName(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile)}/MicrosoftTranslatorAccountkey(x)/value");
            Log?.Info($"put the secret in {Path.GetFileName(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile)}/MicrosoftTranslatorAccountPwd(x)/value");
        }

        /// <summary>
        /// Initialises  the languages supported
        /// </summary>
        /// <returns></returns>
        private static async Task Initlanguages()
        {
            try
            {
                _codeToLanguage = new Dictionary<string, string>();
                _languageToCode = new Dictionary<string, string>();
                var code = await HttpGet<List<string>>(_baseUri + "/GetLanguagesForTranslate");
                var names = await HttpPost<List<string>>(_baseUri + "/GetLanguageNames?locale=en", code);
                for (int i = 0; i < code.Count; i++)
                {
                    _codeToLanguage.Add(code[i], names[i]);
                    _languageToCode.Add(names[i], code[i]);
                }
            }
            catch(WebException e)
            {
                Log?.Error(e.Message);
            }
        }



        #endregion
    }
}
