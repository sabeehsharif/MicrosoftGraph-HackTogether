
using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Linq;
using System.Net;
using System.Net.Http;

namespace DotNetCoreRazor_MSGraph.Graph
{
    public class GraphEmailClient
    {
        private readonly ILogger<GraphEmailClient> _logger = null;
        private readonly GraphServiceClient _graphServiceClient = null;

        public GraphEmailClient(ILogger<GraphEmailClient> logger, GraphServiceClient graphServiceClient)
        {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }

        public async Task<IEnumerable<Message>> GetUserMessages()
        {
            try
            {
//                List<Option> options = new List<Option>
//{
//   //Creating header
//   new HeaderOption("Prefer","outlook.body-content-type='text'")
//};
                //var test = _graphServiceClient.Me.MailFolders.Inbox.Messages;
                //var emails = await _graphServiceClient.Me.Messages
                var emails = await _graphServiceClient.Me.MailFolders.Inbox.Messages
            .Request()
            .Select(msg => new
            {
                msg.Subject,
                msg.Body,
                msg.ReceivedDateTime,
                msg.Id,
                msg.InternetMessageId,
                msg.ConversationId
            })
            .OrderBy("receivedDateTime desc")
            .Top(10)
            .GetAsync();
                return emails.CurrentPage;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error calling Graph /me/messages: {ex.Message}");
                throw;
            }
            // Remove this code
            //return await Task.FromResult<IEnumerable<Message>>(null);
        }
        public async Task<string> GetUserMessageDetails(string id)
        {
            try
            {
                //                List options = new List
                //{

                //new HeaderOption("Prefer", "outlook.body-content-type='text'")
                //};
                List<Option> options = new List<Option>
{
   //Creating header
   new HeaderOption("Prefer","outlook.body-content-type='text'")
};
                var emails = await _graphServiceClient.Me.Messages[id].Request(options).Select(msg => new
            {
                    msg.Subject,
                msg.BodyPreview,
                msg.ReceivedDateTime,
                msg.Id,
                msg.InternetMessageId,
                msg.ConversationId,
                msg.Body

                }).GetAsync();
                //.Request()
                //.Select(msg => new
                //{
                //    msg.Subject,
                //    msg.BodyPreview,
                //    msg.ReceivedDateTime,
                //    msg.Id,
                //    msg.InternetMessageId,
                //    msg.ConversationId
                //})
                //.OrderBy("receivedDateTime desc")
                //.Top(10)
                //.GetAsync();
                return emails.Body.Content;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error calling Graph /me/messages: {ex.Message}");
                throw;
            }
            // Remove this code
            //return await Task.FromResult<IEnumerable<Message>>(null);
        }
        public async Task<(IEnumerable<Message> Messages, string NextLink)> GetUserMessagesPage(
            string nextPageLink = null, int top = 10)
        {
            // Remove this code
            return await Task.FromResult<
                (IEnumerable<Message> Messages, string NextLink)>((Messages:null, NextLink:null));
        }

        private string GetNextLink(IUserMessagesCollectionPage pagedMessages) {
            if (pagedMessages.NextPageRequest != null)
            {
                // Get the URL for the next batch of records
                return pagedMessages.NextPageRequest.GetHttpRequestMessage().RequestUri?.OriginalString;
            }
            return null;
        }

    }
}