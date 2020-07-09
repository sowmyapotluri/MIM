// <copyright file="BookAMeetingBot.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Bots
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using AdaptiveCards;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Dialogs;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Connector;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Teams.Apps.Bart.Cards;
    using Microsoft.Teams.Apps.Bart.Helpers;
    using Microsoft.Teams.Apps.Bart.Models;
    using Microsoft.Teams.Apps.Bart.Models.TableEntities;
    using Microsoft.Teams.Apps.Bart.Providers.Interfaces;
    using Microsoft.Teams.Apps.Bart.Providers.Storage;
    using Microsoft.Teams.Apps.Bart.Resources;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Implements the core logic of the Book A Room bot.
    /// </summary>
    /// <typeparam name="T">Generic class.</typeparam>
    public class BartBot<T> : TeamsActivityHandler
        where T : Dialog
    {
        /// <summary>
        /// Reads and writes conversation state for your bot to storage.
        /// </summary>
        private readonly BotState conversationState;

        /// <summary>
        /// Dialog to be invoked.
        /// </summary>
        private readonly Dialog dialog;

        /// <summary>
        /// Stores user specific data.
        /// </summary>
        private readonly BotState userState;

        /// <summary>
        /// Application base URL.
        /// </summary>
        private readonly string appBaseUri;

        /// <summary>
        /// Instrumentation key for application insights logging.
        /// </summary>
        private readonly string instrumentationKey;

        /// <summary>
        /// Valid tenant id for which bot will operate.
        /// </summary>
        private readonly string tenantId;

        /// <summary>
        /// Generating and validating JWT token.
        /// </summary>
        private readonly ITokenHelper tokenHelper;

        /// <summary>
        /// Storage provider to perform insert, update and delete operation on ActivityEntities table.
        /// </summary>
        private readonly IActivityStorageProvider activityStorageProvider;

        /// <summary>
        /// Storage provider to perform insert, update and delete operation on UserFavorites table.
        /// </summary>
        private readonly IServiceNowProvider serviceNowProvider;

        /// <summary>
        /// Telemetry client for logging events and errors.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Storage provider to perform insert and update operation on UserConfiguration table.
        /// </summary>
        private readonly IUserConfigurationStorageProvider userConfigurationStorageProvider;

        /// <summary>
        /// Helper class which exposes methods required for incident creation.
        /// </summary>
        private readonly IConferenceBridgesStorageProvider conferenceBridgesStorageProvider;

        private readonly MicrosoftAppCredentials microsoftAppCredentials;

        private readonly IIncidentStorageProvider incidentStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="BartBot{T}"/> class.
        /// </summary>
        /// <param name="conversationState">Reads and writes conversation state for your bot to storage.</param>
        /// <param name="userState">Reads and writes user specific data to storage.</param>
        /// <param name="dialog">Dialog to be invoked.</param>
        /// <param name="tokenHelper">Generating and validating JWT token.</param>
        /// <param name="activityStorageProvider">Storage provider to perform insert, update and delete operation on ActivityEntities table.</param>
        /// <param name="serviceNowProvider">Provider for exposing methods required to perform meeting creation.</param>
        /// <param name="telemetryClient">Telemetry client for logging events and errors.</param>
        /// <param name="userConfigurationStorageProvider">Storage provider to perform insert and update operation on UserConfiguration table.</param>
        /// <param name="appBaseUri">Application base URL.</param>
        /// <param name="instrumentationKey">Instrumentation key for application insights logging.</param>
        /// <param name="tenantId">Valid tenant id for which bot will operate.</param>
        public BartBot(ConversationState conversationState, UserState userState, T dialog, ITokenHelper tokenHelper, IActivityStorageProvider activityStorageProvider, IServiceNowProvider serviceNowProvider, TelemetryClient telemetryClient, IUserConfigurationStorageProvider userConfigurationStorageProvider, IIncidentStorageProvider incidentStorageProvider, string appBaseUri, string instrumentationKey, string tenantId, MicrosoftAppCredentials microsoftAppCredentials, IConferenceBridgesStorageProvider conferenceBridgesStorageProvider)
        {
            this.conversationState = conversationState;
            this.userState = userState;
            this.dialog = dialog;
            this.tokenHelper = tokenHelper;
            this.activityStorageProvider = activityStorageProvider;
            this.serviceNowProvider = serviceNowProvider;
            this.telemetryClient = telemetryClient;
            this.userConfigurationStorageProvider = userConfigurationStorageProvider;
            this.appBaseUri = appBaseUri;
            this.instrumentationKey = instrumentationKey;
            this.tenantId = tenantId;
            this.microsoftAppCredentials = microsoftAppCredentials;
            this.incidentStorageProvider = incidentStorageProvider;
            this.conferenceBridgesStorageProvider = conferenceBridgesStorageProvider;
        }

        /// <summary>
        /// Handles an incoming activity.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public override async Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken = default)
        {
            var activity = turnContext.Activity;
            if (!this.IsActivityFromExpectedTenant(turnContext))
            {
                this.telemetryClient.TrackTrace($"Unexpected tenant Id {activity.Conversation.TenantId}", SeverityLevel.Warning);
                await turnContext.SendActivityAsync(MessageFactory.Text(Strings.InvalidTenant)).ConfigureAwait(false);
            }
            else
            {
                this.telemetryClient.TrackTrace($"Activity received = Activity Id: {activity.Id}, Activity type: {activity.Type}, Activity text: {activity?.Text}, From Id: {activity.From?.Id}, User object Id: {activity.From?.AadObjectId}", SeverityLevel.Information);
                var locale = activity.Entities?.Where(entity => entity.Type == "clientInfo").First().Properties["locale"].ToString();

                // Get the current culture info to use in resource files
                if (locale != null)
                {
                    CultureInfo.CurrentUICulture = CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo(locale);
                }
                this.telemetryClient.TrackTrace($"TurnContext- {turnContext}", SeverityLevel.Information);
                await base.OnTurnAsync(turnContext, cancellationToken).ConfigureAwait(false);

                // Save any state changes that might have occured during the turn.
                await this.conversationState.SaveChangesAsync(turnContext, false, cancellationToken).ConfigureAwait(false);
                await this.userState.SaveChangesAsync(turnContext, false, cancellationToken).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Invoked when members other than this bot (like a user) are added to the conversation.
        /// </summary>
        /// <param name="membersAdded">List of members added.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            var activity = turnContext?.Activity;
            this.telemetryClient.TrackTrace($"conversationType: {activity.Conversation.ConversationType}, membersAdded: {activity.MembersAdded?.Count}, membersRemoved: {activity.MembersRemoved?.Count}");

            if (activity.MembersAdded?.Where(member => member.Id != activity.Recipient.Id).FirstOrDefault() != null)
            {
                this.telemetryClient.TrackEvent("Bot installed", new Dictionary<string, string>() { { "User", activity.From.AadObjectId } });
                var userStateAccessors = this.userState.CreateProperty<UserData>(nameof(UserData));
                var userdata = await userStateAccessors.GetAsync(turnContext, () => new UserData()).ConfigureAwait(false);

                if (userdata?.IsWelcomeCardSent == null || userdata?.IsWelcomeCardSent == false)
                {
                    userdata.IsWelcomeCardSent = true;
                    await userStateAccessors.SetAsync(turnContext, userdata).ConfigureAwait(false);
                    var welcomeCardImageUrl = new Uri(baseUri: new Uri(this.appBaseUri), relativeUri: "/images/welcome.jpg");
                    await turnContext.SendActivityAsync(activity: MessageFactory.Attachment(WelcomeCard.GetWelcomeCardAttachment()), cancellationToken).ConfigureAwait(false);
                }
            }
        }

        /// <summary>
        /// Invoked when members other than this bot (like a user) are removed from the conversation.
        /// </summary>
        /// <param name="membersRemoved">List of members removed.</param>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnMembersRemovedAsync(IList<ChannelAccount> membersRemoved, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            var activity = turnContext?.Activity;
            this.telemetryClient.TrackTrace($"conversationType: {activity.Conversation?.ConversationType}, membersAdded: {activity.MembersAdded?.Count}, membersRemoved: {activity.MembersRemoved?.Count}");

            if (activity.MembersAdded?.Where(member => member.Id != activity.Recipient.Id).FirstOrDefault() != null)
            {
                this.telemetryClient.TrackEvent("Bot uninstalled", new Dictionary<string, string>() { { "User", activity.From.AadObjectId } });
                var userStateAccessors = this.userState.CreateProperty<UserData>(nameof(UserData));
                var userdata = await userStateAccessors.GetAsync(turnContext, () => new UserData()).ConfigureAwait(false);
                userdata.IsWelcomeCardSent = false;
                await userStateAccessors.SetAsync(turnContext, userdata).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Invoked when a signin or verify activity is received.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnTeamsSigninVerifyStateAsync(ITurnContext<IInvokeActivity> turnContext, CancellationToken cancellationToken)
        {
            if (turnContext.Activity.Name.Equals("signin/verifyState", StringComparison.OrdinalIgnoreCase))
            {
                await turnContext.SendActivityAsync(MessageFactory.Text(Strings.LoggedInSuccess), cancellationToken).ConfigureAwait(false);
            }

            await this.dialog.RunAsync(turnContext, this.conversationState.CreateProperty<DialogState>(nameof(DialogState)), cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Invoked when task module fetch event is received from the user.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleFetchAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            var activity = turnContext.Activity;
            if (taskModuleRequest.Data == null)
            {
                this.telemetryClient.TrackTrace("Request data obtained on task module fetch action is null.");
                await turnContext.SendActivityAsync(Strings.ExceptionResponse).ConfigureAwait(false);
                return default;
            }

            var userToken = await this.tokenHelper.GetUserTokenAsync(activity.From.Id);
            if (string.IsNullOrEmpty(userToken))
            {
                // No token found for user. Trying to open task module after sign out.
                this.telemetryClient.TrackTrace("User token is null in OnTeamsTaskModuleFetchAsync.");
                await turnContext.SendActivityAsync(Strings.SignInErrorMessage).ConfigureAwait(false);
                return default;
            }

            var postedValues = JsonConvert.DeserializeObject<Data>(JObject.Parse(taskModuleRequest.Data.ToString()).SelectToken("data").ToString());
            string command = postedValues.Text;
            var token = this.tokenHelper.GenerateAPIAuthToken(activity.From.AadObjectId, activity.ServiceUrl, activity.From.Id, jwtExpiryMinutes: 60);
            string activityReferenceId = string.Empty;

            switch (command.ToUpperInvariant())
            {
                // Show task module to manage favorite rooms which is invoked from 'Favorite rooms' list card.
                case BotCommands.CreateIncident:
                    // activityReferenceId = postedValues.ActivityReferenceId;
                    this.telemetryClient.TrackTrace("Create incident executed.");
                    return this.GetTaskModuleResponse(string.Format(CultureInfo.InvariantCulture, "{0}/incident?telemetry={1}&token={2}", this.appBaseUri, this.instrumentationKey, token), Strings.AddFavTaskModuleSubtitle, "large", "large");

                case BotCommands.EditWorkstream:
                    this.telemetryClient.TrackTrace("Edit workstream executed.");
                    activityReferenceId = postedValues.ActivityReferenceId;
                    return this.GetTaskModuleResponse(string.Format(CultureInfo.InvariantCulture, "{0}/editWorkstream?telemetry={1}&token={2}&incident={3}", this.appBaseUri, this.instrumentationKey, token, activityReferenceId), Strings.EditWorkstream, "460", "600");

                default:
                    var reply = MessageFactory.Text(Strings.CommandNotRecognized.Replace("{command}", command, StringComparison.OrdinalIgnoreCase));
                    await turnContext.SendActivityAsync(reply).ConfigureAwait(false);
                    return default;
            }
        }

        /// <summary>
        /// Invoked when task module submit event is received from the user.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="taskModuleRequest">Task module invoke request value payload.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task<TaskModuleResponse> OnTeamsTaskModuleSubmitAsync(ITurnContext<IInvokeActivity> turnContext, TaskModuleRequest taskModuleRequest, CancellationToken cancellationToken)
        {
            var activity = turnContext.Activity;
            if (taskModuleRequest.Data == null)
            {
                this.telemetryClient.TrackTrace("Request data obtained on task module submit action is null.");
                await turnContext.SendActivityAsync(Strings.ExceptionResponse).ConfigureAwait(false);
                return default;
            }

            // Not needed
            //var userConfiguration = await this.userConfigurationStorageProvider.GetAsync(activity.From.AadObjectId).ConfigureAwait(false);
            //if (userConfiguration == null)
            //{
            //    this.telemetryClient.TrackTrace("User configuration is null in task module submit action.");
            //    await turnContext.SendActivityAsync(Strings.ExceptionResponse).ConfigureAwait(false);
            //    return default;
            //}
            this.telemetryClient.TrackTrace($"TaskModule Submit- {taskModuleRequest.Data.ToString()}");
            var valuesFromTaskModule = JsonConvert.DeserializeObject<Incident>(taskModuleRequest.Data.ToString());
            if (valuesFromTaskModule != null)
            {
                //IncidentEntity incidentEntity = new IncidentEntity();
                //incidentEntity.PartitionKey = valuesFromTaskModule.Number;
                //var activityId = turnContext.Activity.Id;
                //var cId = turnContext.Activity.Conversation.Id;
                //var messgae = await turnContext.SendActivityAsync(activity: MessageFactory.Attachment(IncidentCard.GetIncidentAttachment(valuesFromTaskModule, false)), cancellationToken).ConfigureAwait(false);
                //await turnContext.SendActivityAsync(activity: MessageFactory.Attachment(IncidentCard.TestCard(JsonConvert.SerializeObject(messgae))), cancellationToken).ConfigureAwait(false);

                //var attachment = new Attachment
                //{
                //    ContentType = AdaptiveCard.ContentType,
                //    Content = IncidentCard.TestCard(JsonConvert.SerializeObject(messgae)),
                //};

                //var updateCardActivity = new Activity(ActivityTypes.Message)
                //{
                //    Id = activityId,
                //    Conversation = new ConversationAccount { Id = cId },
                //    Attachments = new List<Attachment> { attachment },
                //};
                //await turnContext.UpdateActivityAsync(updateCardActivity, cancellationToken).ConfigureAwait(false);

                var personalChatActivityId = await turnContext.SendActivityAsync(activity: MessageFactory.Attachment(IncidentCard.GetIncidentAttachment(valuesFromTaskModule)), cancellationToken).ConfigureAwait(false);
                ConversationResourceResponse teamCard = await this.SendCardToTeamAsync(turnContext, IncidentCard.GetIncidentAttachment(valuesFromTaskModule), "19:7295b1ef36c64bf2a3052f02103b240c@thread.tacv2", cancellationToken).ConfigureAwait(false);
                IncidentEntity incidentEntity = new IncidentEntity();
                incidentEntity.PartitionKey = valuesFromTaskModule.Number;
                incidentEntity.TeamConversationId = teamCard.Id;
                incidentEntity.TeamActivityId = teamCard.ActivityId;
                incidentEntity.ServiceUrl = teamCard.ServiceUrl;
                incidentEntity.RowKey = valuesFromTaskModule.Id;
                incidentEntity.PersonalConversationId = turnContext.Activity.Conversation.Id;
                incidentEntity.PersonalActivityId = personalChatActivityId.Id;
                var insert = await this.incidentStorageProvider.AddAsync(incidentEntity).ConfigureAwait(false);
            }
            else
            {
                await turnContext.SendActivityAsync(activity: MessageFactory.Attachment(WelcomeCard.GetWelcomeCardAttachment()), cancellationToken).ConfigureAwait(false);
            }
            //var message = valuesFromTaskModule.Text;
            //var replyToId = valuesFromTaskModule.ReplyTo;

            //if (message.Equals(BotCommands.MeetingFromTaskModule, StringComparison.OrdinalIgnoreCase))
            //{
            //    var attachment = SuccessCard.GetSuccessAttachment(valuesFromTaskModule, userConfiguration.WindowsTimezone);
            //    var activityFromStorage = await this.activityStorageProvider.GetAsync(activity.From.AadObjectId, replyToId).ConfigureAwait(false);

            //    if (!string.IsNullOrEmpty(replyToId))
            //    {
            //        var updateCardActivity = new Activity(ActivityTypes.Message)
            //        {
            //            Id = activityFromStorage.ActivityId,
            //            Conversation = activity.Conversation,
            //            Attachments = new List<Attachment> { attachment },
            //        };
            //        await turnContext.UpdateActivityAsync(updateCardActivity).ConfigureAwait(false);
            //    }

            //    await turnContext.SendActivityAsync(MessageFactory.Text(string.Format(CultureInfo.CurrentCulture, Strings.RoomBooked, valuesFromTaskModule.RoomName)), cancellationToken).ConfigureAwait(false);
            //    }
            //    else
            //    {
            //        if (!string.IsNullOrEmpty(replyToId))
            //        {
            //            await this.UpdateFavouriteCardAsync(turnContext, replyToId).ConfigureAwait(false);
            //}
            //    }

            return null;
        }

        /// <summary>
        /// Invoked when a message activity is received from the user.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var command = turnContext.Activity.Text;
            await this.SendTypingIndicatorAsync(turnContext).ConfigureAwait(false);

            if (turnContext.Activity.Text == null && turnContext.Activity.Value != null && turnContext.Activity.Type == ActivityTypes.Message
                && (!string.IsNullOrEmpty(JToken.Parse(turnContext.Activity.Value.ToString()).SelectToken("Action").ToString())))
            {
                command = "CardAction";
            }

            switch (command.ToUpperInvariant())
            {
                case BotCommands.Help:
                    break;
                case "UPDATEACTIVITY":



                    break;

                case "CARDACTION":
                    var values = JsonConvert.DeserializeObject<ChangeTicketStatusPayload>(turnContext.Activity.Value.ToString());
                    var inc = await this.incidentStorageProvider.GetAsync(values.IncidentNumber, values.IncidentId).ConfigureAwait(false);
                    var updateObject = new Incident();
                    updateObject.Id = values.IncidentId;
                    updateObject.Status = "2";
                    updateObject = await this.serviceNowProvider.UpdateIncidentAsync(updateObject, "U1ZDX3RlYW1zX2F1dG9tYXRpb246eWV0KTVUajgmSjkhQUFa").ConfigureAwait(false);
                    updateObject.BridgeDetails = await this.conferenceBridgesStorageProvider.GetAsync("711752242").ConfigureAwait(false);
                    if (turnContext.Activity.Conversation.ConversationType.ToLower() == "teams")
                    {
                        var attachment = new Attachment
                        {
                            ContentType = AdaptiveCard.ContentType,
                            Content = IncidentCard.GetIncidentAttachment(updateObject, "Incident close", false),
                        };

                        var updateCardActivity1 = new Activity(ActivityTypes.Message)
                        {
                            Id = inc.PersonalActivityId,
                            Conversation = new ConversationAccount { Id = inc.PersonalConversationId },
                            Attachments = new List<Attachment> { attachment },
                        };
                        var connector = new ConnectorClient(new Uri(inc.ServiceUrl), this.microsoftAppCredentials);
                        var updateResponse = await connector.Conversations.UpdateActivityAsync(inc.PersonalConversationId, inc.PersonalActivityId, updateCardActivity1, cancellationToken); //.UpdateActivityAsync(updateCardActivity, cancellationToken).ConfigureAwait(false);

                    }
                    else
                    {
                        var attachment1 = new Attachment
                        {
                            ContentType = AdaptiveCard.ContentType,
                            Content = IncidentCard.GetIncidentAttachment(updateObject),
                        };
                        var attachment = new Attachment
                        {
                            ContentType = AdaptiveCard.ContentType,
                            Content = IncidentCard.GetIncidentAttachment(updateObject, "Incident close", false),
                        };
                        //var updateCardActivity = new Activity(ActivityTypes.Message)
                        //{
                        //    Id = turnContext.Activity.Id,
                        //    Conversation = new ConversationAccount { Id = turnContext.Activity.Conversation.Id},
                        //    Attachments = new List<Attachment> { attachment1 },
                        //};

                        //var updateCardActivity1 = new Activity(ActivityTypes.Message)
                        //{
                        //    Id = inc.PersonalActivityId,
                        //    Conversation = new ConversationAccount { Id = inc.PersonalConversationId },
                        //    Attachments = new List<Attachment> { attachment },
                        //};
                        //await turnContext.SendActivityAsync(activity: MessageFactory.Attachment(IncidentCard.GetIncidentAttachment(updateObject, "Incident close", false))).ConfigureAwait(false);
                        var connector = new ConnectorClient(new Uri(inc.ServiceUrl), this.microsoftAppCredentials);
                        //var updateResponse = await turnContext.UpdateActivityAsync(updateCardActivity1, cancellationToken); //.UpdateActivityAsync(updateCardActivity, cancellationToken).ConfigureAwait(false);
                        var updateResponse = await connector.Conversations.UpdateActivityAsync(inc.PersonalConversationId, inc.PersonalActivityId, (Activity)MessageFactory.Attachment(IncidentCard.GetIncidentAttachment(updateObject)), cancellationToken).ConfigureAwait(false); //.UpdateActivityAsync(updateCardActivity, cancellationToken).ConfigureAwait(false);
                    }
                    break;

                default:
                    await this.dialog.RunAsync(turnContext, this.conversationState.CreateProperty<DialogState>(nameof(DialogState)), cancellationToken).ConfigureAwait(false);
                    break;
            }
        }

        /// <summary>
        /// Send help card containing commands recognized by bot.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        //private static async Task ShowHelpCardAsync(ITurnContext<IMessageActivity> turnContext)
        //{
        //    var activity = (Activity)turnContext.Activity;
        //    var reply = activity.CreateReply();
        //    reply.Attachments = HelpCard.GetHelpAttachments();
        //    reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;
        //    await turnContext.SendActivityAsync(reply).ConfigureAwait(false);
        //}

        /// <summary>
        /// Send typing indicator to the user.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task SendTypingIndicatorAsync(ITurnContext turnContext)
        {
            try
            {
                var typingActivity = turnContext.Activity.CreateReply();
                typingActivity.Type = ActivityTypes.Typing;
                await turnContext.SendActivityAsync(typingActivity);
            }
            catch (Exception ex)
            {
                // Do not fail on errors sending the typing indicator
                this.telemetryClient.TrackTrace($"{ex} Failed to send a typing indicator");
            }
        }

        /// <summary>
        /// Send the given attachment to the specified team.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="cardToSend">The card to send.</param>
        /// <param name="teamId">Team id to which the message is being sent.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns><see cref="Task"/>That resolves to a <see cref="ConversationResourceResponse"/>Send a attachemnt.</returns>
        private async Task<ConversationResourceResponse> SendCardToTeamAsync(
            ITurnContext turnContext,
            Attachment cardToSend,
            string teamId,
            CancellationToken cancellationToken)
        {
            var conversationParameters = new ConversationParameters
            {
                Activity = (Activity)MessageFactory.Attachment(cardToSend),
                ChannelData = new TeamsChannelData { Channel = new ChannelInfo(teamId) },
            };
            var taskCompletionSource = new TaskCompletionSource<ConversationResourceResponse>();
            await ((BotFrameworkAdapter)turnContext.Adapter).CreateConversationAsync(
                null,       // If we set channel = "msteams", there is an error as preinstalled middleware expects ChannelData to be present.
                turnContext.Activity.ServiceUrl,
                this.microsoftAppCredentials,
                conversationParameters,
                (newTurnContext, newCancellationToken) =>
                {
                    var activity = newTurnContext.Activity;
                    taskCompletionSource.SetResult(new ConversationResourceResponse
                    {
                        Id = activity.Conversation.Id,
                        ActivityId = activity.Id,
                        ServiceUrl = activity.ServiceUrl,
                    });
                    return Task.CompletedTask;
                },
                cancellationToken).ConfigureAwait(false);
            return await taskCompletionSource.Task.ConfigureAwait(false);
        }


        /// <summary>
        /// Verify if the tenant Id in the message is the same tenant Id used when application was configured.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <returns>Boolean indicating whether tenant is valid.</returns>
        private bool IsActivityFromExpectedTenant(ITurnContext turnContext)
        {
            return turnContext.Activity.Conversation.TenantId.Equals(this.tenantId, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Get task module response object.
        /// </summary>
        /// <param name="url">Task module URL.</param>
        /// <param name="title">Title for task module.</param>
        /// <param name="height">Task module height.</param>
        /// <param name="width">Task module width.</param>
        /// <returns>TaskModuleResponse object.</returns>
        private TaskModuleResponse GetTaskModuleResponse(string url, string title, string height, string width)
        {
            return new TaskModuleResponse
            {
                Task = new TaskModuleContinueResponse
                {
                    Type = "continue",
                    Value = new TaskModuleTaskInfo()
                    {
                        Url = url,
                        Height = "large",
                        Width = "large",
                        Title = title,
                    },
                },
            };
        }
    }
}
