using log4net;
using Microsoft.Crm.UnifiedServiceDesk.Dynamics.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization;
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
            /// <param name="key">The key.</param>
            /// <param name="pwd">The password.</param>
            public   CloudIdentifier(string key, string pwd)
            {
                Key = key;
                Password = pwd;
            }
        }

        #region private
        private static int _index = 0;
        //manages till 4 differents accounts. to be set in unitycomments.dll.config
        private static List<CloudIdentifier> CloudIds = new List<CloudIdentifier>()
        {
            new CloudIdentifier( Properties.Settings.Default.MicrosoftTranslatorAccountkey1,Properties.Settings.Default.MicrosoftTranslatorAccountPwd1),
            new CloudIdentifier( Properties.Settings.Default.MicrosoftTranslatorAccountkey2,Properties.Settings.Default.MicrosoftTranslatorAccountPwd2),
            new CloudIdentifier( Properties.Settings.Default.MicrosoftTranslatorAccountkey3,Properties.Settings.Default.MicrosoftTranslatorAccountPwd3),
            new CloudIdentifier( Properties.Settings.Default.MicrosoftTranslatorAccountkey4,Properties.Settings.Default.MicrosoftTranslatorAccountPwd4)
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
            return await HttpGet<string>(BuildDetectUri(text));
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
                    string Responsebuilder = string.Empty;
                    foreach (string s in parts)
                    {
                        Responsebuilder += await HttpGet<string>(BuildTranslateUri(text, from, to));
                    }
                    translation = Responsebuilder;
                    if (translation != null)
                    {
                        Log?.Debug($"Translation for source text '{text}' from {from} to {to} is {translation}");
                    }
                    else
                    {
                        Log?.Warn($"Translation for source text '{text}' from {from} to {to} does not return a valid string");
                        translation = text;//back to original;
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
        /// <returns></returns>
        private static async Task<T> HttpGet<T>(string url, int retry = 3)
        {
            for (int i = 0; i < retry; i++)
            {
                try
                {
                    Initialize();
                    HttpWebRequest request = CreateRequest(url);
                    return await GetResponse<T>(request);
                }
                catch(Exception e)
                {
                    if (retry == i + 1)
                    {
                        //last try
                        Log?.Error(e.Message);
                    }
                    CheckCredentials(true);
                    _credential = null; //forces the recalculation of the token
                }
            }
            return default(T);
        }


        /// <summary>
        /// HTTP post.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="uri">The URI.</param>
        /// <param name="param">The parameter.</param>
        /// <returns></returns>
        private static async Task<T> HttpPost<T>(string uri, object param = null, int retry =3)
        {
            for (int i = 0; i < retry; i++)
            {
                try
                {
                    Initialize();
                    // create the request
                    HttpWebRequest request = CreateRequest(uri);
                    SetPostParameters(param, request);
                    return await GetResponse<T>(request);
                }
                catch (Exception e)
                {
                    if (retry == i + 1)
                    {
                        //last try
                        Log?.Error(e.Message);
                    }
                    CheckCredentials(true);
                    _credential = null; //forces the recalculation of the token
                }
            }
            return default(T);
        }

        /// <summary>
        /// Gets the response.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="request">The request.</param>
        /// <returns></returns>
        private static async Task<T> GetResponse<T>(HttpWebRequest request)
        {
            WebResponse response = null;
            using (response = await request.GetResponseAsync())
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
                DataContractSerializer dcs = new DataContractSerializer(typeof(T));//(Type.GetType("System.String"));
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

                CheckCredentials();
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
                    Log?.Error(e);
                }

            }
        }

        /// <summary>
        /// Checks the credentials.
        /// </summary>
        /// <param name="forceIncr">if set to <c>true</c> [force incr].</param>
        private static void CheckCredentials(bool forceIncr = false)
        {
            if (forceIncr)
            {
                _index = (_index + 1) % CloudIds.Count;
            }
            if (!CloudID.Valid)
            {
                int start = _index;
                while (!CloudID.Valid)
                {
                    if (!FindNextCredentials())
                    {
                        Log?.Error($"No credential founds for Microsoft translator - please update unitycomments.dll.config with valid credentials");
                        Log?.Info("Get Client Id and Client Secret from https://datamarket.azure.com/developer/applications");
                        Log?.Info("Refer obtaining AccessToken (http://msdn.microsoft.com/en-us/library/hh454950.aspx)");
                        Log?.Info("this tool support till 4 microsoft accounts");
                        Log?.Info("put the client ID in unitycomments.dll.config/MicrosoftTranslatorAccountkey{1-4}/value");
                        Log?.Info("put the secret in unitycomments.dll.config/MicrosoftTranslatorAccountPwd{1-4}/value");
                    }

                }
            }
        }

        /// <summary>
        /// Initialises  the languages supported
        /// </summary>
        /// <returns></returns>
        private static async Task Initlanguages()
        {
            var code = await HttpGet<List<string>>(_baseUri + "/GetLanguagesForTranslate");
            var names = await HttpPost<List<string>>(_baseUri + "/GetLanguageNames?locale=en", code);
            _codeToLanguage = new Dictionary<string, string>();
            _languageToCode = new Dictionary<string, string>();
            for (int i = 0; i < code.Count; i++)
            {
                _codeToLanguage.Add(code[i], names[i]);
                _languageToCode.Add(names[i], code[i]);
            }
        }



        #endregion
    }
}
