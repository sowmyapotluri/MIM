// <copyright file="MessagingExtensionTicketsCard.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Cards
{
    using System;
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.CodeAnalysis.CSharp.Syntax;
    using Microsoft.Teams.Apps.Bart.Cards;
    using Microsoft.Teams.Apps.Bart.Models;
    using Microsoft.Teams.Apps.Bart.Models.TableEntities;


    /// <summary>
    /// Implements messaging extension tickets card.
    /// </summary>
    public class MessagingExtenstionCard : IncidentCard
    {
        private IncidentEntity incident = new IncidentEntity();

        /// <summary>
        /// Initializes a new instance of the <see cref="MessagingExtenstionCard"/> class.
        /// </summary>
        /// <param name="incidentEntity">The incident entity in table storage.</param>
        /// <param name="incident">The incident model with the latest details.</param>
        public MessagingExtenstionCard(IncidentEntity incidentEntity, Incident incident)
            : base(incident)
        {
            this.incident = incidentEntity;
        }

        /// <summary>
        /// Return the appropriate set of card actions based on the state and information in the ticket.
        /// </summary>
        /// <returns>Adaptive card actions.</returns>
        protected override List<AdaptiveAction> BuildActions()
        {
            List<AdaptiveAction> actions = new List<AdaptiveAction>();

            if (!string.IsNullOrEmpty(this.incident.TeamConversationId))
            {
                actions.Add(
                    new AdaptiveOpenUrlAction
                    {
                        Title = "Go to original thread",
                        Url = new Uri(CreateDeeplinkToThread(this.incident.TeamConversationId)),
                    });
            }

            return actions;
        }

        /// <summary>
        /// Returns go to original thread uri which will help in opening the original conversation about the incident.
        /// </summary>
        /// <param name="threadConversationId">The thread along with message Id stored in storage table.</param>
        /// <returns>Original thread uri.</returns>
        private static string CreateDeeplinkToThread(string threadConversationId)
        {
            string[] threadAndMessageId = threadConversationId.Split(";");
            var threadId = threadAndMessageId[0];
            var messageId = threadAndMessageId[1].Split("=")[1];
            return $"https://teams.microsoft.com/l/message/{threadId}/{messageId}";
        }
    }
}
