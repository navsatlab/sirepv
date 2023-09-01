using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.Apis.Services;
using Google.Apis;
using Google.Apis.Forms.v1;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;
using System.Net.Http;

namespace TSI.SIREP.Modules
{
    internal class GForms
    {
        #region Declares
        private readonly string _jsonpath = AppDomain.CurrentDomain.BaseDirectory + @"credential.json";
        private string _idForm, _ApiKey, _urlForm;

        FormsService _svcForm;
        UserCredential _cred;
        #endregion

        public GForms()
        {
            
        }

        #region Methods
        public async Task<bool> InitializeLoad()
        {
            _idForm = await GetFormId(_urlForm);

            var _secret = await GoogleClientSecrets.FromFileAsync(_jsonpath);
            _cred = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                _secret.Secrets,
                new[] { FormsService.ScopeConstants.FormsBodyReadonly },
                "SIREPV_OAUTH", CancellationToken.None);
            
            string apiURL = $"https://forms.googleapis.com/v1/forms/{_idForm}";
            var client = new HttpClient();
            var response = await client.GetAsync(apiURL);
            var content = await response.Content.ReadAsStringAsync();

            var _svcInit = new BaseClientService.Initializer
            {
                HttpClientInitializer = _cred,
                ApplicationName = "SIREPV_OAUTH",
                ApiKey = _ApiKey
            };
            _svcForm = new FormsService(_svcInit);


            try
            {
                var form = await _svcForm.Forms.Get(_idForm).ExecuteAsync();

                var questionIds = new List<string>();

                if (form.Items != null && form.Items.Count > 0)
                {
                    foreach (var question in form.Items)
                    {
                        questionIds.Add(question.ItemId);
                    }
                }


            }
            catch { }


            return true;
        }
        #endregion

        #region Parameters
        /// <summary>
        /// Obtiene o establece la ruta de acceso web del formulario de Google
        /// </summary>
        public string UrlForm
        {
            get { return _urlForm; }
            set { _urlForm = value; }
        }

        /// <summary>
        /// Obtiene o establece la Api del servicio de Google Forms
        /// </summary>
        public string ApiKey
        {
            get { return _ApiKey; }
            set { _ApiKey = value; }
        }
        #endregion

        #region PrivateMethods
        private async Task<string> GetFormId(string formUrl)
        {
            using (var client = new HttpClient())
            {
                var response = await client.GetAsync(formUrl);
                var content = await response.Content.ReadAsStringAsync();

                // Parse the form ID from the HTML content
                var formIdStartIndex = content.IndexOf(@"docs-crp"":""/forms/d/e/") + 22;
                var formIdEndIndex = content.IndexOf("/viewform", formIdStartIndex);
                var formId = content.Substring(formIdStartIndex, formIdEndIndex - formIdStartIndex);

                return formId;
            }
        }

        private async Task<string[]> GetQuestionsIds(BaseClientService.Initializer _var)
        {
            using (var service = new FormsService(_var))
            {
                var form = await service.Forms.Get(_idForm).ExecuteAsync();

                var questionIds = new List<string>();

                if (form.Items != null && form.Items.Count > 0)
                {
                    foreach (var question in form.Items)
                    {
                        questionIds.Add(question.ItemId);
                    }
                }

                return questionIds.ToArray();
            }
        }
        #endregion
    }
}
