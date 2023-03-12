using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Azure;
using Azure.AI.TextAnalytics;
using DotNetCoreRazor.Graph;
using DotNetCoreRazor_MSGraph.CognitiveService;
using DotNetCoreRazor_MSGraph.Graph;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Web;

namespace DotNetCoreRazor_MSGraph.Pages
{
    [AuthorizeForScopes(ScopeKeySection = "DownstreamApi:Scopes")]
    public class EmailModel : PageModel
    {
        private readonly GraphEmailClient _graphEmailClient;
        private readonly GraphTeamsClient _graphTeamsClient;

        [BindProperty(SupportsGet = true)]
        public string NextLink { get; set; }
        public IEnumerable<Message> Messages { get; private set; }
        public List<string> SummarizedTextResult;
        public EmailModel(GraphEmailClient graphEmailClient, GraphTeamsClient graphTeamsClient)
        {
            _graphEmailClient = graphEmailClient;
            _graphTeamsClient = graphTeamsClient;
        }

        public async Task OnGetAsync()
        {
            Messages = await _graphEmailClient.GetUserMessages();
            //string MessageId = "AAMkADZkMzY5ZWVlLTFkNzEtNGMwYi05NDQ3LWVjNjJkNjIzNmFiNQBGAAAAAACdscMaReZSQ7oF3KPwwSZxBwDwHLYZ0j_qTreja8fBp_umAAAAAAEPAADwHLYZ0j_qTreja8fBp_umAAIasSz0AAA=";
            //var selectedUserMessage = await _graphEmailClient.GetUserMessageDetails(MessageId);
            //string emailBody = @"The extractive summarization feature uses natural language processing techniques to locate key sentences in an unstructured text document. 
            //        These sentences collectively convey the main idea of the document. This feature is provided as an API for developers. 
            //        They can use it to build intelligent solutions based on the relevant information extracted to support various use cases. 
            //        In the public preview, extractive summarization supports several languages. It is based on pretrained multilingual transformer models, part of our quest for holistic representations. 
            //        It draws its strength from transfer learning across monolingual and harness the shared nature of languages to produce models of improved quality and efficiency.";
            //var SummarizedTextResult = await CognitiveServiceSummarization.GenerateSummarizedText(emailBody);
            //SummarizedTextResult = await CognitiveServiceSummarization.GenerateSummarizedText(emailBody);


            //var client = new TextAnalyticsClient(endpoint, credentials);
            //AnalyzeActionsOperation = await CognitiveServiceSummarization.TextSummarizationExample(client);
        }
        //public async Task<IActionResult> OnGetAsyncUpdateSearchResults(string selectedMessageId)
        //{
        //    //int[] types = selectedTypes.Split(",").Select(x => int.Parse(x)).ToArray();

        //    //var inventory = await _itemService.GetFiltered(types, null, null, null, null, null, null, startDate, endDate.ToUniversalTime(), null, null, null, null, null, null, null);

        //    //if (inventory != null)
        //    //{
        //    //    SearchResultsGridPartialModel = new SearchResultsGridPartialModel();
        //    //    SearchResultsGridPartialModel.TotalCount = inventory.TotalCount;
        //    //    SearchResultsGridPartialModel.TotalPages = inventory.TotalPages;
        //    //    SearchResultsGridPartialModel.PageNumber = inventory.PageNumber;
        //    //    SearchResultsGridPartialModel.Items = inventory.Items;
        //    //}
        //    //string MessageId = "AAMkADZkMzY5ZWVlLTFkNzEtNGMwYi05NDQ3LWVjNjJkNjIzNmFiNQBGAAAAAACdscMaReZSQ7oF3KPwwSZxBwDwHLYZ0j_qTreja8fBp_umAAAAAAEPAADwHLYZ0j_qTreja8fBp_umAAIasSz0AAA=";
        //    var selectedUserMessage = await _graphEmailClient.GetUserMessageDetails(selectedMessageId);
        //    MessageViewModel message = new MessageViewModel();
        //    //message.BodyPreview = selectedUserMessage;
        //    // var SummarizedTextResult = await CognitiveServiceSummarization.GenerateSummarizedText(selectedUserMessage);
        //    //message.Body = SummarizedTextResult.ToString();
        //    string summarizedEmail = await SummarizedEmail(selectedUserMessage);
        //    message.Body = summarizedEmail;
        //    var myViewData = new ViewDataDictionary(new Microsoft.AspNetCore.Mvc.ModelBinding.EmptyModelMetadataProvider(), new Microsoft.AspNetCore.Mvc.ModelBinding.ModelStateDictionary()) { { "SearchResultsGridPartialModel", message.Body } };
        //    myViewData.Model = message;

        //    PartialViewResult result = new PartialViewResult()
        //    {
        //        ViewName = "SummarizedText",
        //        ViewData = myViewData,
        //    };

        //    return result;
        //}
        public async Task<IActionResult> OnGetAsyncGetTeamsChannel(string selectedMessageId, string TeamsId)
        {
            var selectedUserMessage = await _graphEmailClient.GetUserMessageDetails(selectedMessageId);
            MessageViewModel message = new MessageViewModel();
            message.Body = null;
            //message.Body = selectedUserMessage;
            var summarizedEmailPoints = await SummarizedEmail(selectedUserMessage);
            //List<string> summarizedEmail = new List<string>();
            //foreach (var item in summarizedEmailPoints)
            //{
            //    summarizedEmail.Add(item.ToString());
            //}
            message.Body = summarizedEmailPoints;

            var channelsList = await _graphTeamsClient.GetTeamsChannels(TeamsId);
            message.selectedChannel = channelsList.FirstOrDefault().Id;
            //var responsePostedMessage = await PostEmailToChannel(TeamsId , channelsList.FirstOrDefault().Id, summarizedEmailPoints);

            Dictionary<string, string> teamsChannelsList = new Dictionary<string, string>();
            foreach (var item in channelsList)
            {
                teamsChannelsList.Add(item.DisplayName, item.Id);
            }
            message.channelsList = teamsChannelsList;

            var myViewData = new ViewDataDictionary(new Microsoft.AspNetCore.Mvc.ModelBinding.EmptyModelMetadataProvider(), new Microsoft.AspNetCore.Mvc.ModelBinding.ModelStateDictionary()) { { "SearchResultsGridPartialModel", message } };
            myViewData.Model = message;

            PartialViewResult result = new PartialViewResult()
            {
                ViewName = "SummarizedText",
                ViewData = myViewData,
            };

            return result;
        }
        public async Task<IEnumerable> SummarizedEmail(string selectedUserMessage)
        {
            var SummarizedTextResult = await CognitiveServiceSummarization.GenerateSummarizedText(selectedUserMessage);
            return SummarizedTextResult;
        }
        public async Task<bool> PostEmailToChannel(string TeamsId, string ChannelId, IEnumerable EmailBody)
        {
            string summarizedBodyForTeams = "";
            int i = 1;
            foreach (var item in EmailBody)
            {
                summarizedBodyForTeams = summarizedBodyForTeams + i + ": " + item.ToString();
                i++;
            }
            var ResponsePostedMessage = await _graphTeamsClient.SendMessageToTeamsChannels(TeamsId,ChannelId, summarizedBodyForTeams);
            return ResponsePostedMessage;
        }
        public IActionResult ShowPartailView()
        {
            MessageViewModel message = new MessageViewModel();
            message.Body = "test sfs";
            return Partial("~/Pages/SummarizedText.cshtml", message);
        }
        public PartialViewResult OnPostGetDetails(string emailId)
        {
            string emailId1 = emailId;
            //return new PartialViewResult
            //{
            //    ViewName = "Details",
            //    //ViewData = new ViewDataDictionary<Customer>(ViewData, this.Context.Customers.Find(customerId))
            //    //ViewData = new ViewDataDictionary<string>(ViewData, "test")
            //    ViewData = new ViewDataDictionary<string>(ViewData, await _graphEmailClient.GetUserMessageDetails(emailId1))

            //};
            return new PartialViewResult
            {
                ViewName = "SummarizedText",
                //ViewData = new ViewDataDictionary<Customer>(ViewData, this.Context.Customers.Find(customerId))
                //ViewData = new ViewDataDictionary<string>(ViewData, "test")
                ViewData = new ViewDataDictionary<string>(ViewData, "test")

            };
        }
    }
    class MessageViewModel
    {
        //public string Body { get; set; }
        public IEnumerable Body { get; set; }
        public Dictionary<string, string> channelsList = new Dictionary<string, string>();
        public string selectedChannel { get; set; }
    }
}

