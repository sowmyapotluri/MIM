// <copyright file="IncidentCard.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Cards
{
    using System;
    using System.Collections.Generic;
    using AdaptiveCards;
    using Microsoft.AspNetCore.Mvc.Formatters.Internal;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.CodeAnalysis.CSharp.Syntax;
    using Microsoft.EntityFrameworkCore.Migrations;
    using Microsoft.Teams.Apps.Bart.Models;
    using Microsoft.Teams.Apps.Bart.Models.Enum;
    using Microsoft.Teams.Apps.Bart.Models.TableEntities;
    using Microsoft.Teams.Apps.Bart.Resources;
    using Newtonsoft.Json;

    /// <summary>
    /// Class having method to return incident card attachment.
    /// </summary>
    public class IncidentCard
    {

        private readonly Incident incident = new Incident();

        /// <summary>
        /// Initializes a new instance of the <see cref="IncidentCard"/> class.
        /// </summary>
        /// <param name="incident">Incident object.</param>
        public IncidentCard(Incident incident)
        {
            this.incident = incident;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="IncidentCard"/> class.
        /// </summary>
        public IncidentCard()
        {
        }

        /// <summary>
        /// Get welcome card attachment.
        /// </summary>
        /// <param name="incidentEntity">Incident object from table storage.</param>
        /// <param name="title">Title text for the card.</param>
        /// <param name="footer">Flag to show the status of the incident.</param>
        /// <returns>Adaptive card attachment for bot introduction and bot commands to start with.</returns>
        public Attachment GetIncidentAttachment(IncidentEntity incidentEntity = null, string title = "New incident reported", bool footer = false)
        {
            var footerContainer = new AdaptiveContainer();
            var activityColumnSet = new AdaptiveColumnSet
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
                                Width = "auto",
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
                                                    IncidentId = this.incident.Id,
                                                    IncidentNumber = this.incident.Number,
                                                    Text = "UpdateActivity",
                                                },
                                            },
                                        },
                                    },
                                },
                            },
                        },
            };
            var scopeColumnSet = new AdaptiveColumnSet();
            var bridgeColumnSet = new AdaptiveColumnSet();

            if (footer || this.incident.Status != "1")
            {
                string footerMessage = title != "New incident reported" ? "closed" : this.incident.Status == "2" ? "suspended" : "service restored";
                footerContainer = new AdaptiveContainer
                {
                    Style = AdaptiveContainerStyle.Attention,
                    Bleed = true,
                    Items = new List<AdaptiveElement>
                    {
                        new AdaptiveTextBlock
                        {
                            Weight = AdaptiveTextWeight.Bolder,
                            Text = $"* Please do not respond to this incident, as it is {footerMessage}",
                            Wrap = true,
                        },
                    },
                };

                activityColumnSet = new AdaptiveColumnSet();
            }

            if (!string.IsNullOrEmpty(this.incident.Scope))
            {
                scopeColumnSet = new AdaptiveColumnSet
                {
                    Spacing = AdaptiveSpacing.Medium,
                    Columns = new List<AdaptiveColumn> {
                            new AdaptiveColumn
                            {
                                Width = "150px",
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = Strings.Scope,
                                        Weight = AdaptiveTextWeight.Bolder,
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
                                        Text = this.incident.Scope,
                                        Wrap = true,
                                        Spacing = AdaptiveSpacing.None,
                                    },
                                },
                            },
                        },
                };
            }

            if (this.incident.BridgeDetails.Code != null && this.incident.BridgeDetails.Code != "0")
            {
                bridgeColumnSet = new AdaptiveColumnSet
                {
                    Spacing = AdaptiveSpacing.Medium,
                    Columns = new List<AdaptiveColumn> {
                            new AdaptiveColumn
                            {
                                Width = "150px",
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = Strings.Bridge,
                                        Weight = AdaptiveTextWeight.Bolder,
                                        Wrap = true,
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
                                        Text = string.Format("[{0}]({1})", this.incident.BridgeDetails.Code, this.incident.BridgeDetails.BridgeURL),
                                        Wrap = true,
                                        Spacing = AdaptiveSpacing.None,
                                    },
                                },
                            },
                        },
                };
            }

            AdaptiveCard card = new AdaptiveCard("1.2")
            {
                Body = new List<AdaptiveElement>
                {
                    new AdaptiveContainer
                    {
                        Style = title == "New incident reported" ? AdaptiveContainerStyle.Warning : AdaptiveContainerStyle.Good,
                        Bleed = true,
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
                                               Color = this.incident.Priority == "7" ? AdaptiveTextColor.Attention : AdaptiveTextColor.Default,
                                               HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                               Text = string.Format("{0} priority!", Enum.GetName(typeof(Priority), Convert.ToInt32(this.incident.Priority))),
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
                                        Text = string.Format("ID# {0}", this.incident.Number),
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
                                        Color = this.incident.Status == "1" ? AdaptiveTextColor.Good: AdaptiveTextColor.Default,
                                        HorizontalAlignment = AdaptiveHorizontalAlignment.Right,
                                        Text = this.incident.Status == "1" ? Strings.NewStatusText : this.incident.Status == "2" ? Strings.SuspendedStatusText : Strings.ServiceRestoredStatusText,
                                    },
                                },
                            },
                        },
                    },
                    new AdaptiveColumnSet {
                        Columns = new List<AdaptiveColumn> {
                            new AdaptiveColumn
                            {
                                Width = "150px",
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = Strings.CreatedOn,
                                        Weight = AdaptiveTextWeight.Bolder,
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
                                        Text = this.incident.CreatedOn,
                                        Wrap = true,
                                        Spacing = AdaptiveSpacing.None,
                                    },
                                },
                            },
                        },
                    },
                    scopeColumnSet,
                    new AdaptiveColumnSet
                    {
                        Spacing = AdaptiveSpacing.Medium,
                        Columns = new List<AdaptiveColumn> {
                            new AdaptiveColumn
                            {
                                Width = "150px",
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = Strings.Description,
                                        Weight = AdaptiveTextWeight.Bolder,
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
                                        Text = this.incident.Description,
                                        Wrap = true,
                                        Spacing = AdaptiveSpacing.None,
                                    },
                                },
                            },
                        },
                    },
                    activityColumnSet,
                    new AdaptiveColumnSet
                    {
                        Spacing = AdaptiveSpacing.Medium,
                        Columns = new List<AdaptiveColumn> {
                            new AdaptiveColumn
                            {
                                Width = "150px",
                                Items = new List<AdaptiveElement>
                                {
                                    new AdaptiveTextBlock
                                    {
                                        Text = Strings.ShortDescription,
                                        Weight = AdaptiveTextWeight.Bolder,
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
                                        Text = this.incident.Short_Description,
                                        Wrap = true,
                                        Spacing = AdaptiveSpacing.None,
                                    },
                                },
                            },
                        },
                    },
                    bridgeColumnSet,
                    footerContainer,
                },
                Actions = !footer ? this.BuildActions() : new List<AdaptiveAction>(),
            };
            var adaptiveCardAttachment = new Attachment()
            {
                ContentType = AdaptiveCard.ContentType,
                Content = card,
            };

            return adaptiveCardAttachment;
        }

        /// <summary>
        /// Return the appropriate set of card actions based on the state and information in the incident.
        /// </summary>
        /// <returns>Adaptive card actions.</returns>
        protected virtual List<AdaptiveAction> BuildActions()
        {
            return new List<AdaptiveAction>
                {
                    new AdaptiveSubmitAction
                    {
                        Title = "View workstreams",
                        Data = new AdaptiveSubmitActionData
                        {
                            Msteams = new TaskModuleAction(Strings.ViewWorkstreams, new { data = JsonConvert.SerializeObject(new AdaptiveTaskModuleCardAction { Text = BotCommands.EditWorkstream, ActivityReferenceId = this.incident.Id, ActivityReferenceNumber = this.incident.Number }) }),
                        },
                    },
                    new AdaptiveShowCardAction
                    {
                        Title = "Change Status",
                        Card = new AdaptiveCard(new AdaptiveSchemaVersion(1, 0))
                        {
                            Body = new List<AdaptiveElement>
                            {
                                new AdaptiveColumnSet
                                {
                                    Columns = new List<AdaptiveColumn>
                                    {
                                        new AdaptiveColumn
                                        {
                                           Items = new List<AdaptiveElement>
                                           {
                                                GetAdaptiveChoiceSetTitleInput(),
                                           },
                                        },
                                        new AdaptiveColumn
                                        {
                                           Items = new List<AdaptiveElement>
                                           {
                                                GetAdaptiveChoiceSetStatusInput(this.incident),
                                           },
                                        },
                                    },
                                },
                            },
                            Actions = new List<AdaptiveAction>
                            {
                                new AdaptiveSubmitAction
                                {
                                    Data = new ChangeTicketStatusPayload { IncidentId = this.incident.Id, IncidentNumber = this.incident.Number },
                                },
                            },
                        },
                    },
                };
        }

        /// <summary>
        /// Return the appropriate fact set based on the state and information in the ticket.
        /// </summary>
        /// <param name="incident">Incident object.</param>
        /// <param name="half">Flag identifier to know which factset set.</param>
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
                if (!string.IsNullOrEmpty(incident.Scope))
                {
                    factList.Add(new AdaptiveFact
                    {
                        Title = "Scope",
                        Value = incident.Scope,
                    });
                }

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
                if (incident.BridgeDetails.Code != null && incident.BridgeDetails.Code != "0")
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
        private static AdaptiveChoiceSetInput GetAdaptiveChoiceSetStatusInput(Incident incident)
        {
            AdaptiveChoiceSetInput choiceSet = new AdaptiveChoiceSetInput
            {
                Id = nameof(ChangeTicketStatusPayload.Action),
                IsMultiSelect = false,
                Style = AdaptiveChoiceInputStyle.Compact,
            };

            choiceSet.Value = ChangeTicketStatusPayload.SuspendedAction;
            choiceSet.Choices = new List<AdaptiveChoice>
                    {
                        //new AdaptiveChoice
                        //{
                        //    Title = "New",
                        //    Value = ChangeTicketStatusPayload.NewAction,
                        //},
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
        /// Return the appropriate status choices for incident status.
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

            choiceSet.Value = "Incident Closed";
            choiceSet.Choices = new List<AdaptiveChoice>
                    {
                        new AdaptiveChoice
                        {
                            Title = "Incident Closed",
                            Value = "Incident Closed",
                        },
                    };

            return choiceSet;
        }
    }
}
