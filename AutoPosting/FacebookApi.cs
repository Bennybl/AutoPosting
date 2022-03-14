using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace AutoPosting
{
    public class FacebookApi
    {
        private readonly string FB_PAGE_ID;
        private readonly string FB_ACCESS_TOKEN;
        private const string FB_BASE_ADDRESS = "https://graph.facebook.com/";

        public FacebookApi(string pageId, string accessToken)
        {
            FB_PAGE_ID = pageId;
            FB_ACCESS_TOKEN = accessToken;
        }

        public async Task<string> PublishMessage(string message)
        {
            using (var httpClient = new HttpClient())
            {
                httpClient.BaseAddress = new Uri(FB_BASE_ADDRESS);

                var parametters = new Dictionary<string, string>
                {
                    { "access_token", FB_ACCESS_TOKEN },
                    { "message", message }
                };
                var encodedContent = new FormUrlEncodedContent(parametters);

                var result = await httpClient.PostAsync($"{FB_PAGE_ID}/feed", encodedContent);
                var msg = result.EnsureSuccessStatusCode();
                return await msg.Content.ReadAsStringAsync();
            }

        }
    }
}
