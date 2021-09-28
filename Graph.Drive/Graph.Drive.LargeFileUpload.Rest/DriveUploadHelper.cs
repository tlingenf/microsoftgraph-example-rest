using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace Graph.Drive.LargeFileUpload.Rest
{
    class DriveUploadHelper : IDisposable
    {
        private string _accessToken = null;
        private string _driveUri = null;
        private HttpClient _client = null;

        public DriveUploadHelper(string accessToken, string driveUri)
        {
            if (string.IsNullOrEmpty(accessToken))
                throw new ArgumentNullException("accessToken");

            if (string.IsNullOrEmpty(driveUri))
                throw new ArgumentNullException("driveUri");

            _driveUri = driveUri;

            _client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
            _client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));            
        }

        public string UploadUrl { get; private set; }

        public void Dispose()
        {
            _client.Dispose();
        }

        private async Task<string> CreateUploadSession(string filePath)
        {
            var createSessionResponse = await _client.PostAsync($"{_driveUri}/createUploadSession", new StringContent(""));
            if (createSessionResponse.StatusCode == System.Net.HttpStatusCode.OK)
            {
                dynamic response = JObject.Parse(await createSessionResponse.Content.ReadAsStringAsync());
                this.UploadUrl = response.uploadUrl;
                return this.UploadUrl;
            }
            else
            {
                throw new HttpRequestException("Unable to create upload session.");
            }
        }



        private async Task UploadBytes( int rangeStart)
        {
            await _client.PutAsync(this.UploadUrl, new content)
        }
    }
}
