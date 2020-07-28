// <copyright file="WelcomeCard.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Cards
{
    using System;
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Teams.Apps.Bart.Models;
    using Microsoft.Teams.Apps.Bart.Resources;
    using Newtonsoft.Json;

    /// <summary>
    /// Class having method to return welcome card attachment.
    /// </summary>
    public static class WelcomeCard
    {
        /// <summary>
        /// Get welcome card attachment.
        /// </summary>
        /// <returns>Adaptive card attachment for bot introduction and bot commands to start with.</returns>
        public static Attachment GetWelcomeCardAttachment()
        {
            AdaptiveCard card = new AdaptiveCard("1.0")
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {

                            new AdaptiveColumn
                            {
                                Width = AdaptiveColumnWidth.Auto,
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Large,
                                        Text = Strings.WelcomeText,
                                        Wrap = true,
                                        Weight = AdaptiveTextWeight.Bolder,
                                    },
                                    new AdaptiveTextBlock
                                    {
                                        Text = Strings.WelcomeCardTextLine1,
                                        Wrap = true,
                                    },
                                },
                            },
                        },
                    },

                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = Strings.WelcomeCardTextLine2,
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = Strings.WelcomeCardTextLine3,
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = Strings.WelcomeCardTextLine4,
                        Wrap = true,
                    },
                    new AdaptiveTextBlock
                    {
                        HorizontalAlignment = AdaptiveHorizontalAlignment.Left,
                        Text = Strings.WelcomeCardTextLine5,
                        Wrap = true,
                    },
                },
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = Strings.CreateIncident,
                        Data = new AdaptiveSubmitActionData
                        {
                            Msteams = new TaskModuleAction(Strings.CreateIncident, new { data = JsonConvert.SerializeObject(new AdaptiveTaskModuleCardAction { Text = BotCommands.CreateIncident }) }),
                        },
                    },
                    new AdaptiveSubmitAction
                    {
                        Title = Strings.TakeTour,
                        Data = new TeamsAdaptiveSubmitActionData
                        {
                            MsTeams = new CardAction
                            {
                              Type = ActionTypes.MessageBack,
                              Title = Strings.TakeTour,
                              DisplayText = Strings.TakeTour,
                              Text = BotCommands.TakeTour,
                            },
                        },
                    },
                },
            };

            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };

            return adaptiveCardAttachment;
        }
    }
}
