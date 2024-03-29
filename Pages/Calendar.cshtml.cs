using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DotNetCoreRazor.Graph;
using DotNetCoreRazor_MSGraph.CognitiveService;
using DotNetCoreRazor_MSGraph.Graph;
using DotNetCoreRazor_MSGraph.ResponseModel;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.ExternalConnectors;
using Microsoft.Identity.Web;

namespace DotNetCoreRazor_MSGraph.Pages
{
    [AuthorizeForScopes(ScopeKeySection = "DownstreamApi:Scopes")]
    public class CalendarModel : PageModel
    {
        private readonly ILogger<CalendarModel> _logger;
        private readonly GraphCalendarClient _graphCalendarClient;
        private readonly GraphProfileClient _graphProfileClient;
        private readonly GraphSharePointClient _graphSharePointClient;
        public string Message { get; set; }
        private readonly IConfiguration _configuration;
        private MailboxSettings MailboxSettings { get; set; }

        public IEnumerable<Event> Events  { get; private set; }

        public CalendarModel(ILogger<CalendarModel> logger, GraphCalendarClient graphCalendarClient, GraphProfileClient graphProfileClient, GraphSharePointClient graphSharePointClient, IConfiguration configuration)
        {
            _logger = logger;
            _graphCalendarClient = graphCalendarClient;
            _graphProfileClient = graphProfileClient;
            _graphSharePointClient = graphSharePointClient;
            _configuration = configuration;
        }

        public async Task OnGetAsync()
        {
            
           // var result = await GetSharePointListItems();
            //Messages = await _graphEmailClient.GetUserMessages();
            // Remove this code
            await Task.CompletedTask;
        }
        public async void OnPostSave()
        {
            var result = await GetSharePointListItems();
            ViewData["Message"] = "You clicked Save!";
            Message = "clicked";
        }
        public async Task<IEnumerable> GetSharePointListItems()
        {
            string siteID = _configuration.GetValue<string>("ConfigurationSharePoint:siteid");
            string listId = _configuration.GetValue<string>("ConfigurationSharePoint:listid");
            var listResponse = await _graphSharePointClient.GetSharePointListItems(siteID, listId);
            Dictionary<string, string> groceryItems = new Dictionary<string, string>();
            //var resultCurrentPage = listItems.CurrentPage;
            //foreach (var item in listResponse)
            //{
            //    foreach (var itemField in item.Fields.AdditionalData)
            //    {
            //        groceryItems.Add(itemField.Key + item.Id, itemField.Value.ToString());
            //    }
            //}
            //List<string> eventsList = new List<string>();
            //foreach (var item in listResponse)
            //{
            //    string singleEvent="";
            //    foreach (var itemField in item.Fields.AdditionalData)
            //    {
            //        if (itemField.Key == "Title")
            //        {

            //        }
            //        groceryItems.Add(itemField.Key + item.Id, itemField.Value.ToString());
            //    }
            //}
            //var result = CreateEvent(groceryItems);
            foreach (var item in listResponse)
            {
                var result = CreateEvent(item);
            }
            return listResponse;
        }
        public async Task<IEnumerable<Event>> CreateEvent(ResponseListItems groccery)
        {
            var eventResponse = await _graphCalendarClient.CreateEvent(groccery);

            return eventResponse;
        }
            public string FormatDateTimeTimeZone(DateTimeTimeZone value)
        {
            // Parse the date/time string from Graph into a DateTime
            var graphDatetime = value.DateTime;
            if (DateTime.TryParse(graphDatetime, out DateTime dateTime)) 
            {
                var dateTimeFormat = $"{MailboxSettings.DateFormat} {MailboxSettings.TimeFormat}".Trim();
                if (!String.IsNullOrEmpty(dateTimeFormat)) {
                    return dateTime.ToString(dateTimeFormat);
                }
                else 
                {
                    return $"{dateTime.ToShortDateString()} {dateTime.ToShortTimeString()}";
                }
            }
            else
            {
                return graphDatetime;
            }
        }
    }
}
