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
using DotNetCoreRazor_MSGraph.ResponseModel;

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
        public async Task<IEnumerable<ResponseListItems>> GetSharePointListItems(string siteId, string listId)
        {
            try
            {
                //var test = _graphServiceClient.Me.MailFolders.Inbox.Messages;
                //var emails = await _graphServiceClient.Me.Messages
                //    var listItems = await _graphServiceClient.Sites[siteId].Lists[listId]
                //.Request()
                //.GetAsync();
                //    return listItems.Items;
                //var listsTotal = await _graphServiceClient.Sites[siteId].Lists.Request().GetAsync();

            var queryOptions = new List<QueryOption>()
            {
                new QueryOption("expand", "fields(select=Title,ExpiryDate)")
            };

            var listItems = await _graphServiceClient.Sites[siteId].Lists[listId].Items
            .Request(queryOptions)
            .GetAsync();
                var listResponse = listItems.CurrentPage;
                List<ResponseListItems> lstResponseListItems = new List<ResponseListItems>();
                foreach (var item in listResponse)
                {
                    ResponseListItems responseListItems = new ResponseListItems();
                    foreach (var itemField in item.Fields.AdditionalData)
                    {
                        //groceryItems.Add(itemField.Key + item.Id, itemField.Value.ToString());
                        if (itemField.Key == "Title")
                        {
                            responseListItems.Title = itemField.Value.ToString();
                        }
                        else if (itemField.Key == "ExpiryDate")
                        {
                            responseListItems.ExpiryDate = itemField.Value.ToString();
                        }
                    }
                    lstResponseListItems.Add(responseListItems);
                }
                return lstResponseListItems;

                //ListItem listItem = new ListItem();
                //foreach (var item in listItems.CurrentPage)
                //{
                //    listItem.Fields = item.Fields;
                //}
                //Dictionary<string, string> world = new Dictionary<string, string>();
                //var resultCurrentPage = listItems.CurrentPage;
                //foreach (var item in resultCurrentPage)
                //{
                //    foreach (var itemField in item.Fields.AdditionalData)
                //    {

                //    world.Add(itemField.Key+item.Id, itemField.Value.ToString());
                //    }
                //}
                //for(int i=1; i<=2; i++)
                //{
                //    world.Add(resultCurrentPage.FirstOrDefault().Fields.AdditionalData.Keys.FirstOrDefault(), resultCurrentPage.FirstOrDefault().Fields.AdditionalData.Values.FirstOrDefault().ToString());
                //}
                //return listItems.CurrentPage;

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
