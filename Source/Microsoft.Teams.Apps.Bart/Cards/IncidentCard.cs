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
        public static Attachment GetIncidentAttachment(Incident incident, string title = "New Incident reported", bool update = true)
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
                                               Text = title,
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
                                               Text = incident.Priority == "7" ? "High Priority!" : "Priority",
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
                                        Text = update ? "New" : "Updated",
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
                Actions = new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = "View workstream",
                        Data = new AdaptiveSubmitActionData
                        {
                            Msteams = new TaskModuleAction(Strings.OtherRooms, new { data = JsonConvert.SerializeObject(new AdaptiveTaskModuleCardAction { Text = BotCommands.EditWorkstream, ActivityReferenceId = incident.Number }) }),
                        },
                    },
                    //new AdaptiveSubmitAction
                    //{
                    //    Title = "Change Status",
                    //    Data = new Data { Text = "UPDATEACTIVITY" },
                    //},
                    new AdaptiveShowCardAction
                    {
                        Title = "Change Status",
                        Card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 0))
                        {
                            Body = new List<AdaptiveElement>
                            {
                                GetAdaptiveChoiceSetTitleInput(),
                                GetAdaptiveChoiceSetStatusInput(),
                            },
                            Actions = new List<AdaptiveAction>
                            {
                                new AdaptiveSubmitAction
                                {
                                    Data = new ChangeTicketStatusPayload { IncidentId = incident.Id, IncidentNumber = incident.Number },
                                },
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
                if (incident.BridgeDetails.Code != null)
                {
                    factList.Add(new AdaptiveFact
                    {
                        Title = "Incident conference bridge",
                        Value = string.Format("[{0}]({1})", incident.BridgeDetails.Code, incident.BridgeDetails.BridgeURL),
                    });
                }
            }

            return factList;
        }

        /// <summary>
        /// Return the appropriate status choices for ticket status.
        /// </summary>
        /// <returns>An adaptive element which contains the dropdown choices.</returns>
        private static AdaptiveChoiceSetInput GetAdaptiveChoiceSetStatusInput()
        {
            AdaptiveChoiceSetInput choiceSet = new AdaptiveChoiceSetInput
            {
                Id = nameof(ChangeTicketStatusPayload.Action),
                IsMultiSelect = false,
                Style = AdaptiveChoiceInputStyle.Compact,
            };

            choiceSet.Value = ChangeTicketStatusPayload.NewAction;
            choiceSet.Choices = new List<AdaptiveChoice>
                    {
                        new AdaptiveChoice
                        {
                            Title = "New",
                            Value = ChangeTicketStatusPayload.NewAction,
                        },
                        new AdaptiveChoice
                        {
                            Title = "Suspended",
                            Value = ChangeTicketStatusPayload.SuspendedAction,
                        },
                        new AdaptiveChoice
                        {
                            Title = "Service Restored",
                            Value = ChangeTicketStatusPayload.RestoredAction,
                        },
                    };

            return choiceSet;
        }

        /// <summary>
        /// Return the appropriate status choices for ticket status.
        /// </summary>
        /// <returns>An adaptive element which contains the dropdown choices.</returns>
        private static AdaptiveChoiceSetInput GetAdaptiveChoiceSetTitleInput()
        {
            AdaptiveChoiceSetInput choiceSet = new AdaptiveChoiceSetInput
            {
                Id = nameof(ChangeTicketStatusPayload.Title),
                IsMultiSelect = false,
                Style = AdaptiveChoiceInputStyle.Compact,
            };

            choiceSet.Value = ChangeTicketStatusPayload.NewAction;
            choiceSet.Choices = new List<AdaptiveChoice>
                    {
                        new AdaptiveChoice
                        {
                            Title = "Incident New",
                            Value = "Incident New",
                        },
                        new AdaptiveChoice
                        {
                            Title = "Incident Closed",
                            Value = "Incident Closed",
                        },
                    };

            return choiceSet;
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
