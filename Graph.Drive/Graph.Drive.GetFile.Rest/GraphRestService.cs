using Microsoft.Extensions.Configuration;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace Drive.Graph
{
    public class GraphRestService : HttpClient
    {
        public static IConfigurationRoot _configuration;

        const long DOWNLOAD_BUFFER_SIZE = (long)(5 * 1024 * 1024); // buffer = [Size in MB] * 1024 * 1024; Convert to bytes

        public GraphRestService(IConfigurationRoot configuration, string accessToken): base()
        {
            _configuration = configuration;
            this.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
            //this.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
        }

        public async Task<JObject> GetDriveItemAsync(string itemUri)
        {
            var response = await this.GetAsync(itemUri);
            return JObject.Parse(await response.Content.ReadAsStringAsync());
        }

        public async Task<string> DownloadDriveItemAsync(JObject itemInfo)
        {
            long offset = 0; // cursor location for updating the Range header.
            long chunkSize = DOWNLOAD_BUFFER_SIZE;
            byte[] bytesInStream; // bytes in range returned by chunk download.
            // Get the number of bytes to download. calculate the number of chunks and determine
            // the last chunk size.
            int numberOfChunks = Convert.ToInt32(itemInfo["size"].ToObject<int>() / DOWNLOAD_BUFFER_SIZE);
            // We are incrementing the offset cursor after writing the response stream to a file after each chunk. 
            // Subtracting one since the size is 1 based, and the range is 0 base. There should be a better way to do
            // this but I haven't spent the time on that.
            int lastChunkSize = Convert.ToInt32(itemInfo["size"].ToObject<int>() % DOWNLOAD_BUFFER_SIZE) - numberOfChunks;
            if (lastChunkSize > 0) { numberOfChunks++; }

            // Create a file stream to contain the downloaded file.
            string localPath = Path.Combine(_configuration["downloadsFolder"], itemInfo["name"].ToString());
            using (FileStream fileStream = System.IO.File.Create(localPath))
            {
                for (int i = 0; i < numberOfChunks; i++)
                {
                    // Setup the last chunk to request. This will be called at the end of this loop.
                    if (i == numberOfChunks - 1)
                    {
                        chunkSize = lastChunkSize;
                    }

                    // Create the request message with the download URL and Range header.
                    HttpRequestMessage req = new HttpRequestMessage(HttpMethod.Get, itemInfo["@microsoft.graph.downloadUrl"].ToString());
                    req.Headers.Range = new System.Net.Http.Headers.RangeHeaderValue(offset, chunkSize + offset);
                    var response = await this.SendAsync(req);

                    // watch for eTag changes to indicate the file has changed.
                    IEnumerable<string> eTag;
                    response.Headers.TryGetValues("ETag", out eTag);
                    if (eTag.First() != itemInfo["eTag"].ToString())
                    {
                        // file updated during transfer. throw error.
                        throw new InvalidOperationException("eTag missmatch.");
                    }

                    using (Stream responseStream = await response.Content.ReadAsStreamAsync())
                    {
                        bytesInStream = new byte[responseStream.Length];
                        int read = responseStream.Read(bytesInStream, 0, (int)bytesInStream.Length);
                        fileStream.Write(bytesInStream, 0, read);
                    }

                    offset += chunkSize + 1; // Move the offset cursor to the next chunk.
                }
            }

            return localPath;
        }
    }
}
