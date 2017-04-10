using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using NLog;
using SpreadsheetParser.ConnectWise;

namespace SpreadsheetParser.ConnectWise
{
    public class ConnectWiseService : IConnectWiseService
    {
        private readonly Logger _log;
        private static HttpClient _httpClient;
        ApiSettings _settings = null;

        #region Constructor

        public ConnectWiseService(string companyId, string baseUrl, string siteUrl, string siteSuffix, string publicKey, string privateKey)
        {
            LogManager.ThrowExceptions = true;
            _log = LogManager.GetCurrentClassLogger();
            _settings = new ApiSettings(companyId, baseUrl, siteUrl, siteSuffix, publicKey, privateKey);
            InitHttpClient();
        }

        #endregion Constructor

        #region private methods

        private void InitHttpClient()
        {
            _log.Info("InitConnection");
            _httpClient = new HttpClient { BaseAddress = new Uri(_settings.ApiBaseUri) };
            _httpClient.DefaultRequestHeaders.Accept.Clear();
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            _httpClient.DefaultRequestHeaders.Add("x-cw-usertype", "integrator");

            var credentialsString = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{_settings.CompanyId}+{_settings.PublicKey}:{_settings.PrivateKey}"));
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", credentialsString);
        }
        private async Task<T> AddItem<T>(T item, string uri) where T : class
        {
            _log.Info("Adding " + item.GetType().Name);

            var stringContent = ConvertToStringContent(item);
            var response = await _httpClient.PostAsync(uri, stringContent);

            _log.Info($"Status: {response.StatusCode}\nMessage: {response.ReasonPhrase}");

            if (!response.IsSuccessStatusCode)
                await ProcessException(response);

            var addedItem = await response.Content.ReadAsAsync<T>();
            return addedItem;
        }
        private async Task<List<T>> GetItems<T>(string uri)
        {
            _log.Info("Get items " + uri);
            var response = await _httpClient.GetAsync(uri);
            if (!response.IsSuccessStatusCode)
                await ProcessException(response);
            var items = await response.Content.ReadAsAsync<List<T>>();
            return items;
        }
        private async Task<T> GetItem<T>(string uri)
        {
            _log.Info("Get item " + uri);
            var response = await _httpClient.GetAsync(uri);
            if (!response.IsSuccessStatusCode)
                await ProcessException(response);
            var item = await response.Content.ReadAsAsync<T>();
            return item;
        }
        private async Task<T> UpdateItem<T>(T item, string uri) where T : class
        {
            _log.Info("Updating " + item.GetType().Name);

            var stringContent = ConvertToStringContent(item);
            var response = await _httpClient.PutAsync(uri, stringContent);
            if (!response.IsSuccessStatusCode)
                await ProcessException(response);

            var updatedItem = await response.Content.ReadAsAsync<T>();
            return updatedItem;
        }
        private async Task<T> PatchItem<T>(PatchOperation[] operations, string uri) where T : class
        {
            _log.Info("Patching " + uri);

            var stringContent = ConvertToStringContent(operations);
            var method = new HttpMethod("PATCH");
            var request = new HttpRequestMessage(method, uri) { Content = stringContent };
            //HttpResponseMessage response = await _httpClient.SendAsync(request);
            HttpResponseMessage response = _httpClient.SendAsync(request).Result;

            if (!response.IsSuccessStatusCode)
                await ProcessException(response);

            var updatedItem = response.Content.ReadAsAsync<T>().Result;
            return updatedItem;
        }
        private async Task DeleteItem(string uri)
        {
            _log.Info("Deleting " + uri);
            var response = await _httpClient.DeleteAsync(uri);
            if (!response.IsSuccessStatusCode)
                await ProcessException(response);
        }

        private static async Task ProcessException(HttpResponseMessage response)
        {
            var responseMessage = response.Content.ReadAsAsync<ResponseMessage>().Result;
            var exceptionMessage = new StringBuilder($"Status: {response.StatusCode}\t Code: {responseMessage.code}\t Message: {responseMessage.message}");
            if (responseMessage.errors != null)
                foreach (var error in responseMessage.errors)
                    exceptionMessage.AppendLine($"{error.code}\t{error.message}\t{error.resource}\t{error.field}");
            throw new WebException($"ConnectWise error\n{exceptionMessage}");
        }
        private StringContent ConvertToStringContent(object value)
        {
            var postBody = JsonConvert.SerializeObject(value, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore, DateFormatHandling = DateFormatHandling.IsoDateFormat });
            _log.Info("POST body:");
            _log.Info(postBody);
            var stringContent = new StringContent(postBody, Encoding.UTF8, "application/json");
            return stringContent;
        }

        #endregion

        #region Tickets

        public async Task<List<Ticket>> GetTickets()
        {
            return await GetItems<Ticket>(_settings.TicketsUri);
        }

        public async Task<Ticket> GetTicket(int ticketId)
        {
            return await GetItem<Ticket>(_settings.TicketsUri + "/" + ticketId);
        }

        public async Task<Ticket> AddTicket(Ticket ticket)
        {
            return await AddItem(ticket, _settings.TicketsUri);
        }

        public async Task<Ticket> UpdateTicket(Ticket ticket)
        {
            return await UpdateItem(ticket, _settings.TicketsUri + "/" + ticket.id);
        }

        private async Task<Ticket> PatchTicket(int ticketId, PatchOperation operation)
        {
            return await PatchItem<Ticket>(new[] { operation }, _settings.TicketsUri + "/" + ticketId + "/" + _settings.SuffitxUri);
        }

        public async Task<Ticket> CancelTicket(int ticketId)
        {
            return await PatchTicket(ticketId, PatchOperation.CancelTicket());
        }

        public async Task<Ticket> CloseTicket(int ticketId)
        {
            return await PatchTicket(ticketId, PatchOperation.CloseTicket());
        }

        public async Task<Ticket> ChangeCompany(int ticketId, string companyId)
        {
            return await PatchTicket(ticketId, PatchOperation.ChangeTicket(companyId));
        }

        public async Task<Ticket> ChangeGenerically(int ticketId, string companyId, string operation, string path)
        {
            return await PatchTicket(ticketId, PatchOperation.ChangeGenericOpPath(companyId, operation, path));
        }

        #endregion

        #region Properties       

        #endregion Properties
    }

    #region Helper classes

    internal class ApiSettings
    {
        public string CompanyId { get; }
        public string PublicKey { get; }
        public string PrivateKey { get; }
        public string ApiBaseUri { get; }
        public string TicketsUri { get; }
        public string SuffitxUri { get; }

        public ApiSettings(string companyId, string baseUrl, string siteUrl, string suffixUrl, string publicKey, string privateKey)
        {
            CompanyId = companyId;
            PublicKey = publicKey;
            PrivateKey = privateKey;
            ApiBaseUri = baseUrl;
            TicketsUri = siteUrl;
            SuffitxUri = suffixUrl;
        }
    }

    internal class CwKeys
    {
        [JsonProperty("publicKey")]
        public string PublicKey { get; set; }
        [JsonProperty("privateKey")]
        public string PrivateKey { get; set; }
        [JsonProperty("expiration")]
        public DateTime Expiration { get; set; }
    }

    #endregion Helper classes
}

