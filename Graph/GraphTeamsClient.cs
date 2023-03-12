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
using DotNetCoreRazor_MSGraph.Graph;

namespace DotNetCoreRazor.Graph
{
    public class GraphTeamsClient
{
        private readonly ILogger<GraphTeamsClient> _logger = null;
        private readonly GraphServiceClient _graphServiceClient = null;

        public GraphTeamsClient(ILogger<GraphTeamsClient> logger, GraphServiceClient graphServiceClient)
        {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }

        public async Task<IEnumerable<Channel>> GetTeamsChannels(string TeamsId)
        {
            try
            {
                //var test = _graphServiceClient.Me.MailFolders.Inbox.Messages;
                //var emails = await _graphServiceClient.Me.Messages
                var channels = await _graphServiceClient.Teams[TeamsId].Channels
            .Request()
            .Select(chnl => new
            {
                chnl.DisplayName,
                chnl.Id,
               
            })
            .GetAsync();
                return channels.CurrentPage;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error calling Graph Teams Channels: {ex.Message}");
                throw;
            }
            // Remove this code
            //return await Task.FromResult<IEnumerable<Message>>(null);
        }

        public async Task<bool> SendMessageToTeamsChannels(string TeamsId, string ChannelId, string EmailBody)
        {
            bool iSMessageSent = false;
            try
            {
                var requestBody = new ChatMessage
                {
                    Body = new ItemBody
                    {
                        Content = EmailBody,
                    },
                };
                var postedMessageResponse = await _graphServiceClient.Teams[TeamsId].Channels[ChannelId].Messages
            .Request().AddAsync(requestBody);
                iSMessageSent = true;
                //return postedMessageResponse.ToString();
                //return null;
                return iSMessageSent;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error Graph Posting Message to Teams Channels: {ex.Message}");
                iSMessageSent = false;
                return iSMessageSent;
                throw;
            }
            // Remove this code
            //return await Task.FromResult<IEnumerable<Message>>(null);
        }
    }
}
