
    using System;
    using System.IO;
    using System.Collections;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    using System.Linq;
    using System.Net;
    using TimeZoneConverter;
    using DotNetCoreRazor.Graph;
    using DotNetCoreRazor_MSGraph.ResponseModel;

    namespace DotNetCoreRazor_MSGraph.Graph
    {
        public class GraphCalendarClient
        {
            private readonly ILogger<GraphCalendarClient> _logger = null;
            private readonly GraphServiceClient _graphServiceClient = null;

            public GraphCalendarClient(ILogger<GraphCalendarClient> logger, GraphServiceClient graphServiceClient)
            {
                _logger = logger;
                _graphServiceClient = graphServiceClient;
            }

            public async Task<IEnumerable<Event>> GetEvents(string userTimeZone)
            {
                // Remove this code
                return await Task.FromResult<IEnumerable<Event>>(null);
            }
            public async Task<IEnumerable<Event>> CreateEvent(ResponseListItems groccery)
            {
                try
                {

                    //            var requestBody = new Event
                    //            {
                    //                Subject = "Let's go for lunch",
                    //                Body = new ItemBody
                    //                {
                    //                    ContentType = BodyType.Html,
                    //                    Content = "Does next month work for you?",
                    //                },
                    //                Start = new DateTimeTimeZone
                    //                {
                    //                    DateTime = "2023-12-12T12:00:00",
                    //                    TimeZone = "Pacific Standard Time",
                    //                },
                    //                End = new DateTimeTimeZone
                    //                {
                    //                    DateTime = "2019-10-10T14:00:00",
                    //                    TimeZone = "Pacific Standard Time",
                    //                },
                    //                Location = new Location
                    //                {
                    //                    DisplayName = "Zaigham",
                    //                },
                    //                Attendees = new List<Attendee>
                    //{
                    //    new Attendee
                    //    {
                    //        EmailAddress = new EmailAddress
                    //        {
                    //            Address = "MuhammdSabeeh@gporg2020.onmicrosoft.com",
                    //            Name = "Muhammad Sabeeh Custom",
                    //        },
                    //        Type = AttendeeType.Required,
                    //    },
                    //},
                    //                IsOnlineMeeting = true,
                    //                OnlineMeetingProvider = OnlineMeetingProviderType.TeamsForBusiness,
                    //            };

                    //            //var result = await _graphServiceClient.Me.Calendars["AQMkADZkMzY5ZWUBLTFkNzEtNGMwYi05NDQANy1lYzYyZDYyMzZhYjUARgAAA52xwxpF5lJDugXco-DBJnEHAPActhnSP6pOt6Nrx8Gn66YAAAIBBgAAAPActhnSP6pOt6Nrx8Gn66YAAAI1VAAAAA=="].Events.Request().AddAsync(requestBody);
                    //            var result = await _graphServiceClient.Me.Calendar.Events.Request().AddAsync(requestBody);
                    var requestBodyMapped = new Event
                    {
                        Subject = "Item Name:" + groccery.Title,
                        Body = new ItemBody
                        {
                            ContentType = BodyType.Html,
                            Content = "Expiry of this item: " + groccery.Title + " is at " + groccery.ExpiryDate,
                        },
                        Start = new DateTimeTimeZone
                        {
                            //DateTime = "2023-03-20T12:00:00",
                            DateTime = groccery.ExpiryDate,
                            TimeZone = "Pacific Standard Time",
                        },
                        End = new DateTimeTimeZone
                        {
                            DateTime = groccery.ExpiryDate,
                            TimeZone = "Pacific Standard Time",
                        },
                        Location = new Location
                        {
                            DisplayName = "Bahrain",
                        },
                        Attendees = new List<Attendee>
     {
     new Attendee
     {
     EmailAddress = new EmailAddress
     {
     Address = "MuhammdSabeeh@gporg2020.onmicrosoft.com",
     Name = "Muhammad Sabeeh Custom",
     },
     Type = AttendeeType.Required,
     },
     },
                        IsOnlineMeeting = true,
                        OnlineMeetingProvider = OnlineMeetingProviderType.TeamsForBusiness,
                    };
                    var result = await _graphServiceClient.Me.Calendar.Events.Request().AddAsync(requestBodyMapped);
                    return result.Instances;
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Error calling Calendar List: {ex.Message}");
                    throw;
                }

                // Remove this code
                //return await Task.FromResult<IEnumerable<Event>>(null);
            }
            // Used for timezone settings related to calendar
            public async Task<MailboxSettings> GetUserMailboxSettings()
            {
                try
                {
                    var currentUser = await _graphServiceClient
                        .Me
                        .Request()
                        .Select(u => new
                        {
                            u.MailboxSettings
                        })
                        .GetAsync();

                    return currentUser.MailboxSettings;
                }
                catch (Exception ex)
                {
                    _logger.LogError($"/me Error: {ex.Message}");
                    throw;
                }
            }

            private static DateTime GetUtcStartOfWeekInTimeZone(DateTime today, string timeZoneId)
            {
                // Time zone returned by Graph could be Windows or IANA style
                // .NET Core's FindSystemTimeZoneById needs IANA on Linux/MacOS,
                // and needs Windows style on Windows.
                // TimeZoneConverter can handle this for us
                TimeZoneInfo userTimeZone = TZConvert.GetTimeZoneInfo(timeZoneId);

                // Assumes Sunday as first day of week
                int diff = System.DayOfWeek.Sunday - today.DayOfWeek;

                // create date as unspecified kind
                var unspecifiedStart = DateTime.SpecifyKind(today.AddDays(diff), DateTimeKind.Unspecified);

                // convert to UTC
                return TimeZoneInfo.ConvertTimeToUtc(unspecifiedStart, userTimeZone);
            }

        }
    }