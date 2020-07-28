// <copyright file="TourCard.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Cards
{
    using System.Collections.Generic;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Teams.Apps.Bart.Models;
    using Microsoft.Teams.Apps.Bart.Resources;
    using Newtonsoft.Json;

    /// <summary>
    /// Implements Welcome Tour Carousel card.
    /// </summary>
    public static class TourCard
    {
        /// <summary>
        /// Create incident carousel card.
        /// </summary>
        /// <param name="appBaseUrl">appBaseUrl.</param>
        /// <returns>card.</returns>
        public static Attachment CreateIncidentCard(string appBaseUrl)
        {
            string imageUri = appBaseUrl + "/createIncident.png";

            HeroCard tourCarouselCard = new HeroCard()
            {
                Title = Strings.ReportIncidentHeaderCarousel,
                //Text = string.Format("{0}", Resources.NewRequestCarouselCardText),
                Images = new List<CardImage>()
                {
                    new CardImage(imageUri),
                },
                Buttons = new List<CardAction>()
                {
                    new TaskModuleAction(Strings.CreateIncident, new { data = JsonConvert.SerializeObject(new AdaptiveTaskModuleCardAction { Text = BotCommands.CreateIncident }) }),
                },
            };

            return tourCarouselCard.ToAttachment();
        }

        /// <summary>
        /// View incident carousel card.
        /// </summary>
        /// <param name="appBaseUrl">appBaseUrl.</param>
        /// <returns>card.</returns>
        public static Attachment ViewIncidentCard(string appBaseUrl)
        {
            string imageUri = appBaseUrl + "/updateWorkstream.png";
            HeroCard tourCarouselCard = new HeroCard()
            {
                Title = Strings.UpdateWorkstreamHeaderCarousel,
                Text = string.Format("{0}", Strings.UpdateWorkstreamTextCarousel),
                Images = new List<CardImage>()
                {
                    new CardImage(imageUri),
                },
            };

            return tourCarouselCard.ToAttachment();
        }

        /// <summary>
        /// Close carousel card.
        /// </summary>
        /// <param name="appBaseUrl">appBaseUrl.</param>
        /// <returns>card.</returns>
        public static Attachment CloseCard(string appBaseUrl)
        {
            string imageUri = appBaseUrl + "/closeIncident.png";
            HeroCard tourCarouselCard = new HeroCard()
            {
                Title = Strings.CloseIncidentTextCarousel,
                //Text = string.Format("{0}", Resources.ManageTravelRequestCarouselCardText),
                Images = new List<CardImage>()
                {
                    new CardImage(imageUri),
                },
            };

            return tourCarouselCard.ToAttachment();
        }

        /// <summary>
        /// Create the set of cards that comprise tour carousel.
        /// </summary>
        /// <param name="appBaseUrl">The base URI where the app is hosted.</param>
        /// <returns>The cards that comprise the team tour.</returns>
        public static IEnumerable<Attachment> GetTourCards(string appBaseUrl)
        {
            return new List<Attachment>()
            {
                CreateIncidentCard(appBaseUrl),
                ViewIncidentCard(appBaseUrl),
                CloseCard(appBaseUrl),
            };
        }
    }
}
