// <copyright file="SearchHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Helpers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Resources;
    using System.Threading.Tasks;
    using System.Web;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.CodeAnalysis.CSharp.Syntax;
    using Microsoft.Teams.Apps.Bart.Cards;
    using Microsoft.Teams.Apps.Bart.Models;
    using Microsoft.Teams.Apps.Bart.Models.TableEntities;
    using Microsoft.Teams.Apps.Bart.Providers.Interfaces;
    using Microsoft.Teams.Apps.Bart.Providers.Storage;
    using Newtonsoft.Json;

    /// <summary>
    /// Class that handles the search activities for messaging extension.
    /// </summary>
    public static class SearchHelper
    {
        /// <summary>
        /// New requests command id in the manifest file.
        /// </summary>
        private const string NewCommandId = "newincidents";

        /// <summary>
        ///  Recents command id in the manifest file.
        /// </summary>
        private const string RecentCommandId = "recents";

        /// <summary>
        /// Suspended requests command id in the manifest file.
        /// </summary>
        private const string SuspendedCommandId = "suspendedincidents";

        /// <summary>
        /// Service restored command id in the manifest file.
        /// </summary>
        private const string ServiceRestoredCommandId = "servicerestoredincidents";

        /// <summary>
        /// All incidents command id in the manifest file.
        /// </summary>
        private const string AllCommandId = "allincidents";

        /// <summary>
        /// Get the results from Azure search service and populate the result (card + preview).
        /// </summary>
        /// <param name="query">Query which the user had typed in message extension search.</param>
        /// <param name="commandId">Command id to determine which tab in message extension has been invoked.</param>
        /// <param name="count">Count for pagination.</param>
        /// <param name="skip">Skip for pagination.</param>
        /// <param name="localTimestamp">Local timestamp of the user activity.</param>
        /// <param name="serviceNowProvider">ServiceNow provider.</param>
        /// <param name="incidentStorageProvider">Incident storage provider.</param>
        /// <returns><see cref="Task"/> Returns MessagingExtensionResult which will be used for providing the card.</returns>
        public static async Task<MessagingExtensionResult> GetSearchResultAsync(
            string query,
            string commandId,
            int? count,
            int? skip,
            DateTimeOffset? localTimestamp,
            IServiceNowProvider serviceNowProvider,
            IIncidentStorageProvider incidentStorageProvider)
        {
            MessagingExtensionResult composeExtensionResult = new MessagingExtensionResult
            {
                Type = "result",
                AttachmentLayout = AttachmentLayoutTypes.List,
                Attachments = new List<MessagingExtensionAttachment>(),
            };

            List<Incident> searchServiceResults = new List<Incident>();
            List<IncidentEntity> searchResults = new List<IncidentEntity>();
            query = string.IsNullOrEmpty(query) ? string.Empty : query;
            // commandId should be equal to Id mentioned in Manifest file under composeExtensions section.
            switch (commandId)
            {
                case RecentCommandId:
                    searchServiceResults = await serviceNowProvider.GetIncidentAsync(RecentCommandId, query, "U1ZDX3RlYW1zX2F1dG9tYXRpb246eWV0KTVUajgmSjkhQUFa").ConfigureAwait(false);
                    composeExtensionResult = await GetMessagingExtensionResult(commandId, localTimestamp, searchServiceResults, incidentStorageProvider).ConfigureAwait(false);
                    break;

                case NewCommandId:
                    searchServiceResults = await serviceNowProvider.GetIncidentAsync(NewCommandId, query, "U1ZDX3RlYW1zX2F1dG9tYXRpb246eWV0KTVUajgmSjkhQUFa").ConfigureAwait(false);
                    composeExtensionResult = await GetMessagingExtensionResult(commandId, localTimestamp, searchServiceResults, incidentStorageProvider).ConfigureAwait(false);
                    break;

                case SuspendedCommandId:
                    searchServiceResults = await serviceNowProvider.GetIncidentAsync(SuspendedCommandId, query, "U1ZDX3RlYW1zX2F1dG9tYXRpb246eWV0KTVUajgmSjkhQUFa").ConfigureAwait(false);
                    composeExtensionResult = await GetMessagingExtensionResult(commandId, localTimestamp, searchServiceResults, incidentStorageProvider).ConfigureAwait(false);
                    break;

                case ServiceRestoredCommandId:
                    searchServiceResults = await serviceNowProvider.GetIncidentAsync(ServiceRestoredCommandId, query, "U1ZDX3RlYW1zX2F1dG9tYXRpb246eWV0KTVUajgmSjkhQUFa").ConfigureAwait(false);
                    composeExtensionResult = await GetMessagingExtensionResult(commandId, localTimestamp, searchServiceResults, incidentStorageProvider).ConfigureAwait(false);
                    break;

                case AllCommandId:
                    searchServiceResults = await serviceNowProvider.GetIncidentAsync(AllCommandId, query, "U1ZDX3RlYW1zX2F1dG9tYXRpb246eWV0KTVUajgmSjkhQUFa").ConfigureAwait(false);
                    composeExtensionResult = await GetMessagingExtensionResult(commandId, localTimestamp, searchServiceResults, incidentStorageProvider).ConfigureAwait(false);
                    break;
            }

            return composeExtensionResult;
        }

        /// <summary>
        /// Get populated result to in messaging extension tab.
        /// </summary>
        /// <param name="commandId">Command id to determine which tab in message extension has been invoked.</param>
        /// <param name="localTimestamp">Local timestamp of the user activity.</param>
        /// <param name="searchServiceResults">List of tickets from Azure search service.</param>
        /// <param name="incidentStorageProvider">Incident storage provider.</param>
        /// <returns><see cref="Task"/> Returns MessagingExtensionResult which will be shown in messaging extension tab.</returns>
        public static async Task<MessagingExtensionResult> GetMessagingExtensionResult(
            string commandId,
            DateTimeOffset? localTimestamp,
            IList<Incident> searchServiceResults,
            IIncidentStorageProvider incidentStorageProvider
            )
        {
            MessagingExtensionResult composeExtensionResult = new MessagingExtensionResult
            {
                Type = "result",
                AttachmentLayout = AttachmentLayoutTypes.List,
                Attachments = new List<MessagingExtensionAttachment>(),
            };

            foreach (var incident in searchServiceResults)
            {
                var incidentDetails = await incidentStorageProvider.GetAsync(incident.Number, incident.Id).ConfigureAwait(false);
                if (incidentDetails != null)
                {
                    incident.BridgeDetails = new ConferenceRoomEntity { Code = incidentDetails.BridgeId, BridgeURL = incidentDetails.BridgeLink };

                    ThumbnailCard previewCard = new ThumbnailCard
                    {
                        Title = incident.Number,
                        Text = GetPreviewCardText(incident, commandId, localTimestamp),
                    };
                    //var incidentObject = new Incident
                    //{
                    //    BridgeDetails = new ConferenceRoomEntity
                    //    {
                    //        BridgeURL = incident.BridgeLink,
                    //        Code = incident.BridgeId,
                    //    },
                    //    Description = incident.Description,
                    //    Short_Description = incident.ShortDescription,
                    //    CreatedOn = incident.Timestamp.ToString(),
                    //};
                    var selectedTicketAdaptiveCard = new MessagingExtenstionCard(incidentDetails, incident);
                    composeExtensionResult.Attachments.Add(selectedTicketAdaptiveCard.GetIncidentAttachment(incidentDetails).ToMessagingExtensionAttachment(previewCard.ToAttachment()));
                }
            }

            return composeExtensionResult;
        }

        /// <summary>
        /// Get the text for the preview card for the result.
        /// </summary>
        /// <param name="ticket">Ticket object for ask an expert action.</param>
        /// <param name="commandId">Command id which indicate the action.</param>
        /// <param name="localTimestamp">Local time stamp.</param>
        /// <returns>Command id as string.</returns>
        private static string GetPreviewCardText(Incident ticket, string commandId, DateTimeOffset? localTimestamp)
        {
            //var ticketStatus = commandId != OpenCommandId ? $"<div style='white-space:nowrap'>{HttpUtility.HtmlEncode(Cards.CardHelper.GetTicketDisplayStatusForSme(ticket))}</div>" : string.Empty;
            var cardText = $@"<div>
                                <div style='white-space:nowrap'>
                                        {HttpUtility.HtmlEncode(ticket.Short_Description)}
                                </div> 
                         </div>";
            //HttpUtility.HtmlEncode(CardHelper.GetFormattedDateInUserTimeZone(ticket.DateCreated, localTimestamp))
            return cardText.Trim();
        }
    }
}
