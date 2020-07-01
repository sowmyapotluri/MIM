// <copyright file="IncidentCard.cs" company="Microsoft Corporation">
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
    public class IncidentCard
    {
        /// <summary>
        /// Get welcome card attachment.
        /// </summary>
        /// <returns>Adaptive card attachment for bot introduction and bot commands to start with.</returns>
        public static Attachment GetIncidentAttachment(Incident incident, bool update = false)
        {
            AdaptiveCard card = new AdaptiveCard("1.2")
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveContainer
                    {
                        Style = AdaptiveContainerStyle.Emphasis,
                        Items = new List<AdaptiveElement>
                        {
                            new AdaptiveColumnSet
                            {
                                Columns = new List<AdaptiveColumn>
                                {
                                   new AdaptiveColumn
                                   {
                                       Width = "auto",
                                       Items = new List<AdaptiveElement>
                                       {
                                           new AdaptiveTextBlock
                                           {
                                               Weight = AdaptiveTextWeight.Bolder,
                                               Size = AdaptiveTextSize.Medium,
                                               Text = update ? "Incident closed" : "New Incident reported",
                                           },
                                       },
                                   },
                                   new AdaptiveColumn
                                   {
                                       Width = "stretch",
                                       Items = new List<AdaptiveElement>
                                       {
                                           new AdaptiveTextBlock
                                           {
                                               Weight = AdaptiveTextWeight.Bolder,
                                               Size = AdaptiveTextSize.Medium,
                                               Color = incident.Priority == "7" ? AdaptiveTextColor.Attention: AdaptiveTextColor.Default,
                                               HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                               Text = incident.Priority == "7" ? "High Priority" : "Priority",
                                           },
                                       },
                                   },
                                },
                            },
                        },
                    },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = "auto",
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Medium,
                                        Color = AdaptiveTextColor.Accent,
                                        Weight = AdaptiveTextWeight.Bolder,
                                        Text = string.Format("ID# {0}", incident.Number),
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Width = "stretch",
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Size = AdaptiveTextSize.Default,
                                        Color = AdaptiveTextColor.Good,
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                        Text = "New",
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveFactSet
                    {
                        Facts = BuildFactSet(incident, true),
                    },
                    new AdaptiveColumnSet
                    {
                        Columns = new List<AdaptiveColumn>
                        {
                            new AdaptiveColumn
                            {
                                Width = "stretch",
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextInput
                                    {
                                        Id = "Activity",
                                        Placeholder = "Type current activity",
                                    },
                                },
                            },
                            new AdaptiveColumn
                            {
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveActionSet
                                    {
                                        Id = "Activity",
                                        Actions = new List<AdaptiveAction>
                                        {
                                            new AdaptiveSubmitAction
                                            {
                                                Title = "Update",
                                                Data = new TeamsAdaptiveSubmitActionData
                                                {
                                                    MsTeams = new CardAction
                                                    {
                                                        Type = ActionTypes.MessageBack,
                                                        DisplayText = "Update",
                                                        Text = "UpdateActivity",
                                                    },
                                                    IncidentId = incident.Id,
                                                    IncidentNumber = incident.Number,
                                                    //BridgeId = incident.Bridge.Code,
                                                },
                                            },
                                        },
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveFactSet
                    {
                        Facts = BuildFactSet(incident, false),
                    },
                },
                //Actions = new List<AdaptiveAction>
                //{
                //    new AdaptiveShowCardAction
                //    {
                //        Title = "View workstream",
                //        Card = new AdaptiveCard("1.2")
                //        {
                //            //Actions = AddPrompts(prompts, questionId, userQuestion, 0, 3),
                //        },
                //    },
                //    new AdaptiveShowCardAction
                //    {
                //        Title = "Change Status",
                //        Card = new AdaptiveCard("1.2")
                //        {
                //            //Actions = AddPrompts(prompts, questionId, userQuestion, 0, 3),
                //        },
                //    },
                //},
            };
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };

            return adaptiveCardAttachment;
        }

        /// <summary>
        /// Return the appropriate fact set based on the state and information in the ticket.
        /// </summary>
        /// <param name="localTimestamp">The current timestamp.</param>
        /// <returns>The fact set showing the necessary details.</returns>
        private static List<AdaptiveFact> BuildFactSet(Incident incident, bool half)
        {
            List<AdaptiveFact> factList = new List<AdaptiveFact>();

            if (half)
            {
                factList.Add(new AdaptiveFact
                {
                    Title = "Created On",
                    Value = incident.CreatedOn,
                });
                factList.Add(new AdaptiveFact
                {
                    Title = "Scope",
                    Value = incident.CreatedOn,
                });
                factList.Add(new AdaptiveFact
                {
                    Title = "Description",
                    Value = incident.Description,
                });
            }
            else
            {
                factList.Add(new AdaptiveFact
                {
                    Title = "Short Description",
                    Value = incident.Short_Description,
                });
                if (incident.Bridge != null)
                {
                    factList.Add(new AdaptiveFact
                    {
                        Title = "Incident conference bridge",
                        Value = string.Format("[{0}]({1})", incident.Bridge.Code, incident.Bridge.BridgeURL),
                    });
                }
                factList.Add(new AdaptiveFact
                {
                    Title = "Description",
                    Value = incident.Description,
                });
            }

            return factList;
        }

        public static Attachment TestCard(string contents)
        {
            AdaptiveCard card = new AdaptiveCard("1.2")
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveTextBlock
                    {
                        Text = contents,
                        Wrap = true,
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
