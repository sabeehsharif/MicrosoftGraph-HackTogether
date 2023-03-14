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
    public class GraphSharePointClient
{
        private readonly ILogger<GraphSharePointClient> _logger = null;
        private readonly GraphServiceClient _graphServiceClient = null;

        public GraphSharePointClient(ILogger<GraphSharePointClient> logger, GraphServiceClient graphServiceClient)
        {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }
        public async Task<IEnumerable<ListItem>> GetSharePointListItems(string siteId, string listId)
        {
            try
            {
                //var test = _graphServiceClient.Me.MailFolders.Inbox.Messages;
                //var emails = await _graphServiceClient.Me.Messages
                var listItems = await _graphServiceClient.Sites[siteId].Lists[listId]
            .Request()
            .GetAsync();
                return listItems.Items;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error calling SharePoint List: {ex.Message}");
                throw;
            }
            // Remove this code
            //return await Task.FromResult<IEnumerable<Message>>(null);
        }
    }
}
