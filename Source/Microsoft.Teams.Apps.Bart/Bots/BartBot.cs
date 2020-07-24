// <copyright file="BartBot.cs" company="Microsoft Corporation">
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
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Dialogs;
    using Microsoft.Bot.Builder.Teams;
    using Microsoft.Bot.Connector;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.CodeAnalysis.CSharp.Syntax;
    using Microsoft.Teams.Apps.Bart.Cards;
    using Microsoft.Teams.Apps.Bart.Helpers;
    using Microsoft.Teams.Apps.Bart.Models;
    using Microsoft.Teams.Apps.Bart.Models.TableEntities;
    using Microsoft.Teams.Apps.Bart.Providers.Interfaces;
    using Microsoft.Teams.Apps.Bart.Resources;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Implements the core logic of the BART bot.
    /// </summary>
    /// <typeparam name="T">Generic class.</typeparam>
    public class BartBot<T> : TeamsActivityHandler
        where T : Dialog
    {
        /// <summary>
        /// Search text parameter name in the manifest file.
        /// </summary>
        private const string SearchTextParameterName = "searchText";

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
        /// Storage provider to perform insert and update operation on ServiceNow table.
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
        /// Storage provider to perform insert, update and delete operation on ConferenceBridges table.
        /// </summary>
        private readonly IConferenceBridgesStorageProvider conferenceBridgesStorageProvider;

        /// <summary>
        /// App credentials.
        /// </summary>
        private readonly MicrosoftAppCredentials microsoftAppCredentials;

        /// <summary>
        /// Storage provider to perform insert, update and delete operation on Incidents table.
        /// </summary>
        private readonly IIncidentStorageProvider incidentStorageProvider;

        /// <summary>
        /// Helper class which exposes methods required for workstream creation.
        /// </summary>
        private readonly IWorkstreamStorageProvider workstreamStorageProvider;

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
        /// <param name="incidentStorageProvider">Storage provider to perform insert and update operation on Incidents table.</param>
        /// <param name="appBaseUri">Application base URL.</param>
        /// <param name="instrumentationKey">Instrumentation key for application insights logging.</param>
        /// <param name="tenantId">Valid tenant id for which bot will operate.</param>
        /// <param name="microsoftAppCredentials">App credentials.</param>
        /// <param name="conferenceBridgesStorageProvider">Storage provider to perform insert and update operation on ConferenceBridges table.</param>
        /// <param name="workstreamStorageProvider">Storage provider to perform insert and update operation on Workstreams table.</param>
        public BartBot(ConversationState conversationState, UserState userState, T dialog, ITokenHelper tokenHelper, IActivityStorageProvider activityStorageProvider, 
            IServiceNowProvider serviceNowProvider, TelemetryClient telemetryClient, IUserConfigurationStorageProvider userConfigurationStorageProvider, 
            IIncidentStorageProvider incidentStorageProvider, string appBaseUri, string instrumentationKey, string tenantId, 
            MicrosoftAppCredentials microsoftAppCredentials, IConferenceBridgesStorageProvider conferenceBridgesStorageProvider, IWorkstreamStorageProvider workstreamStorageProvider)
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
            this.workstreamStorageProvider = workstreamStorageProvider;
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
        /// Invoked when user clicks on "Add new incident" button on messaging extension.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="action">Action to be performed.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Response of messaging extension action.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionfetchtaskasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionFetchTaskAsync(
            ITurnContext<IInvokeActivity> turnContext,
            MessagingExtensionAction action,
            CancellationToken cancellationToken)
        {
            try
            {

                var activity = turnContext.Activity;
                var token = this.tokenHelper.GenerateAPIAuthToken(activity.From.AadObjectId, activity.ServiceUrl, activity.From.Id, jwtExpiryMinutes: 60);
                turnContext.Activity.TryGetChannelData<TeamsChannelData>(out var teamsChannelData);
                if (action.CommandId == "viewincident")
                {
                    return await Task.FromResult(new MessagingExtensionActionResponse
                    {
                        Task = new TaskModuleContinueResponse
                        {
                            Value = new TaskModuleTaskInfo
                            {
                                Height = "large",
                                Width = "large",
                                Title = Strings.CreateIncident,
                                Url = string.Format(CultureInfo.InvariantCulture, "{0}/dashboard?telemetry={1}&token={2}", this.appBaseUri, this.instrumentationKey, token),

                            },
                        },
                    });
                }
                string description = string.Empty;  // variable to hold preloaded description if available.
                if (JObject.Parse(activity.Value.ToString()).ContainsKey("messagePayload"))
                {
                    description = JObject.Parse(activity.Value.ToString())["messagePayload"]["body"]["content"].ToString();
                }

                return await Task.FromResult(new MessagingExtensionActionResponse
                {
                    Task = new TaskModuleContinueResponse
                    {
                        Value = new TaskModuleTaskInfo
                        {
                            Height = "large",
                            Width = "large",
                            Title = Strings.CreateIncident,
                            Url = string.Format(CultureInfo.InvariantCulture, "{0}/incident?telemetry={1}&token={2}&description={3}&displayName={4}", this.appBaseUri, this.instrumentationKey, token, description, turnContext.Activity.From.Name),

                        },
                    },
                });

            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
            }

            return default;
        }

        /// <summary>
        /// Invoked when the user submits a create new incident from Messaging Extensions.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="action">Action to be performed.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Response of messaging extension action.</returns>
        /// <remarks>
        /// Reference link: https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionsubmitactionasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<MessagingExtensionActionResponse> OnTeamsMessagingExtensionSubmitActionAsync(
            ITurnContext<IInvokeActivity> turnContext,
            MessagingExtensionAction action,
            CancellationToken cancellationToken)
        {
            try
            {
                var postedObject = ((JObject)turnContext.Activity.Value).GetValue("data", StringComparison.OrdinalIgnoreCase).ToObject<Incident>();

                if (postedObject == null)
                {
                    return default;
                }

                var bridge = await this.conferenceBridgesStorageProvider.GetAsync(postedObject.Bridge).ConfigureAwait(false);
                var userConversationId = await this.userConfigurationStorageProvider.GetAsync(turnContext.Activity.From.AadObjectId).ConfigureAwait(false);
                var connector = new ConnectorClient(new Uri(turnContext.Activity.ServiceUrl), this.microsoftAppCredentials);
                turnContext.Activity.TryGetChannelData<TeamsChannelData>(out var teamsChannelData);
                var card = new IncidentCard(postedObject).GetIncidentAttachment();

                // Sending cards to team and personal chat
                var cardDetailsInTeams = await this.SendCardToTeamAsync(turnContext, card, bridge.ChannelId, cancellationToken).ConfigureAwait(false);
                var cardDetailsInChat = await connector.Conversations.SendToConversationAsync(userConversationId.ConversationId, (Activity)MessageFactory.Attachment(card));

                IncidentEntity incidentEntity = new IncidentEntity {
                    PartitionKey = postedObject.Number,
                    TeamConversationId = cardDetailsInTeams.Id,
                    TeamActivityId = cardDetailsInTeams.ActivityId,
                    ServiceUrl = cardDetailsInTeams.ServiceUrl,
                    RowKey = postedObject.Id,
                    PersonalConversationId = userConversationId.ConversationId,
                    PersonalActivityId = cardDetailsInChat.Id,
                };
                var insert = await this.incidentStorageProvider.AddAsync(incidentEntity).ConfigureAwait(false);
            }
            catch (Exception ex)
            {

                this.telemetryClient.TrackException(ex);
            }

            return default;
        }

        /// <summary>
        /// Invoked when the user opens the messaging extension or searching any content in it.
        /// </summary>
        /// <param name="turnContext">Context object containing information cached for a single turn of conversation with a user.</param>
        /// <param name="query">Contains messaging extension query keywords.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>Messaging extension response object to fill compose extension section.</returns>
        /// <remarks>
        /// https://docs.microsoft.com/en-us/dotnet/api/microsoft.bot.builder.teams.teamsactivityhandler.onteamsmessagingextensionqueryasync?view=botbuilder-dotnet-stable.
        /// </remarks>
        protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionQuery query, CancellationToken cancellationToken)
        {
            var turnContextActivity = turnContext?.Activity;
            try
            {
                turnContextActivity.TryGetChannelData<TeamsChannelData>(out var teamsChannelData);
                var messageExtensionQuery = JsonConvert.DeserializeObject<MessagingExtensionQuery>(turnContextActivity.Value.ToString());
                var searchQuery = this.GetSearchQueryString(messageExtensionQuery);

                return new MessagingExtensionResponse
                {
                    ComposeExtension = await SearchHelper.GetSearchResultAsync(searchQuery, messageExtensionQuery.CommandId, messageExtensionQuery.QueryOptions.Count, messageExtensionQuery.QueryOptions.Skip, turnContextActivity.LocalTimestamp, this.serviceNowProvider, this.incidentStorageProvider).ConfigureAwait(false),
                };

                //string expertTeamId = await this.configurationProvider.GetSavedEntityDetailAsync(ConfigurationEntityTypes.TeamId).ConfigureAwait(false);

                //if (turnContext != null && teamsChannelData?.Team?.Id == expertTeamId && await this.IsMemberOfSmeTeamAsync(turnContext).ConfigureAwait(false))
                //{

                //}

            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
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
                //await this.userConfigurationStorageProvider.AddAsync(new UserConfigurationEntity { UserAdObjectId = activity.From.AadObjectId, ConversationId = activity.Conversation.Id, ServiceUrl = activity.ServiceUrl, TeamsUserId = activity.From.Id });

                if (userdata?.IsWelcomeCardSent == null || userdata?.IsWelcomeCardSent == false)
                {
                    userdata.IsWelcomeCardSent = true;
                    await userStateAccessors.SetAsync(turnContext, userdata).ConfigureAwait(false);
                    await this.userConfigurationStorageProvider.AddAsync(new UserConfigurationEntity { UserAdObjectId = activity.From.AadObjectId, ConversationId = activity.Conversation.Id, ServiceUrl = activity.ServiceUrl, TeamsUserId = activity.From.Id });
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
                await this.userConfigurationStorageProvider.DeleteAsync(new UserConfigurationEntity { UserAdObjectId = activity.From.AadObjectId });
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
            string activityReferenceNumber = string.Empty;

            switch (command.ToUpperInvariant())
            {
                // Show task module to manage favorite rooms which is invoked from 'Favorite rooms' list card.
                case BotCommands.CreateIncident:
                    this.telemetryClient.TrackTrace("Create incident executed.");
                    return this.GetTaskModuleResponse(string.Format(CultureInfo.InvariantCulture, "{0}/incident?telemetry={1}&token={2}&displayName={3}", this.appBaseUri, this.instrumentationKey, token, turnContext.Activity.From.Name), Strings.CreateIncident, "large", "large");

                case BotCommands.EditWorkstream:
                    this.telemetryClient.TrackTrace("Edit workstream executed.");
                    activityReferenceNumber = postedValues.ActivityReferenceNumber;
                    return this.GetTaskModuleResponse(string.Format(CultureInfo.InvariantCulture, "{0}/editWorkstream?telemetry={1}&token={2}&incident={3}&id={4}", this.appBaseUri, this.instrumentationKey, token, activityReferenceNumber, postedValues.ActivityReferenceId), Strings.EditWorkstream, "460", "600");

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
            if (taskModuleRequest.Data == null)
            {
                this.telemetryClient.TrackTrace("Request data obtained on task module submit action is null.");
                await turnContext.SendActivityAsync(Strings.ExceptionResponse).ConfigureAwait(false);
                return default;
            }

            if (((JObject)taskModuleRequest.Data).ContainsKey("output"))
            {
                var parsedObject = JObject.Parse(taskModuleRequest.Data.ToString());
                if (parsedObject["output"].ToObject<bool>())
                {
                    var user = await this.userConfigurationStorageProvider.GetAsync(parsedObject["assignedToId"].ToString()).ConfigureAwait(false);
                    var incidentEntity = await this.incidentStorageProvider.GetAsync(parsedObject["incidentNumber"].ToString()).ConfigureAwait(false);
                    string name = parsedObject["assignedTo"].ToString();
                    var connector = new ConnectorClient(new Uri(incidentEntity.ServiceUrl), this.microsoftAppCredentials);
                    var activity = new Activity(ActivityTypes.Message)
                    {
                        Id = incidentEntity.TeamActivityId,
                        Conversation = new ConversationAccount { Id = incidentEntity.TeamConversationId },
                        Text = "<at>" + name + "</at>, you are assigned this incident.",
                    };
                    if (user != null)
                    {
                        var mentions = new List<Entity>
                        {
                            new Mention
                            {
                                Text = $"<at>{name}</at>",
                                Mentioned = new ChannelAccount()
                                {
                                    Name = name,
                                    Id = user.TeamsUserId,
                                },
                            },
                        };
                        activity.Entities = mentions;
                    }

                    await connector.Conversations.SendToConversationAsync(activity, cancellationToken);
                }

                this.telemetryClient.TrackTrace("Taskmodule submitted from EditWorkstream.");
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
                var channelDetails = await this.conferenceBridgesStorageProvider.GetAsync(valuesFromTaskModule.Bridge).ConfigureAwait(false);
                var card = new IncidentCard(valuesFromTaskModule).GetIncidentAttachment();
                var personalChatActivityId = await turnContext.SendActivityAsync(activity: MessageFactory.Attachment(card), cancellationToken).ConfigureAwait(false);
                ConversationResourceResponse teamCard = await this.SendCardToTeamAsync(turnContext, card, channelDetails.ChannelId, cancellationToken).ConfigureAwait(false);
                IncidentEntity incidentEntity = new IncidentEntity
                {
                    PartitionKey = valuesFromTaskModule.Number,
                    TeamConversationId = teamCard.Id,
                    TeamActivityId = teamCard.ActivityId,
                    ServiceUrl = teamCard.ServiceUrl,
                    RowKey = valuesFromTaskModule.Id,
                    PersonalConversationId = turnContext.Activity.Conversation.Id,
                    PersonalActivityId = personalChatActivityId.Id,
                    ReplyToId = turnContext.Activity.ReplyToId,
                    Status = "1",
                };
                await this.incidentStorageProvider.AddAsync(incidentEntity).ConfigureAwait(false);
                var workstreams = await this.workstreamStorageProvider.GetAllAsync(valuesFromTaskModule.Number).ConfigureAwait(false);
                var sentWorkstreamNotificationTask = new List<Task>();
                foreach (var workstream in workstreams)
                {
                    if (workstream.New)
                    {
                        var user = await this.userConfigurationStorageProvider.GetAsync(workstream.AssignedToId).ConfigureAwait(false);
                        if (user != null)
                        {
                            MicrosoftAppCredentials.TrustServiceUrl(teamCard.ServiceUrl);
                            var connector = new ConnectorClient(new Uri(teamCard.ServiceUrl), this.microsoftAppCredentials);

                            // Sending cards to team and personal chat
                            sentWorkstreamNotificationTask.Add(connector.Conversations.SendToConversationAsync(user.ConversationId, (Activity)MessageFactory.Attachment(card)));
                        }
                    }

                    workstream.New = false;
                    await this.workstreamStorageProvider.AddAsync(workstream).ConfigureAwait(false);
                }

                await Task.WhenAll(sentWorkstreamNotificationTask).ConfigureAwait(false);
            }
            else
            {
                this.telemetryClient.TrackTrace($"TaskModule Submitted values empty");
                await turnContext.SendActivityAsync(activity: MessageFactory.Text("Taskmodule was empty"), cancellationToken).ConfigureAwait(false);
            }

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

                case "CARDACTION":
                    var adaptiveCardSubmitActionData = JsonConvert.DeserializeObject<TeamsAdaptiveSubmitActionData>(turnContext.Activity.Value.ToString());

                    // If it's a activity update action, else it's a status change action.
                    if (!string.IsNullOrEmpty(adaptiveCardSubmitActionData.Activity))
                    {
                        var activityToUpdate = JsonConvert.DeserializeObject<TeamsAdaptiveSubmitActionData>(turnContext.Activity.Value.ToString());
                        var currentActivity = string.Format("{0}: {1}", turnContext.Activity.From.Name, activityToUpdate.Activity);
                        var updateIncident = new Incident
                        {
                            Id = activityToUpdate.IncidentId,
                            CurrentActivity = currentActivity,
                        };
                        await this.serviceNowProvider.UpdateIncidentAsync(updateIncident, "U1ZDX3RlYW1zX2F1dG9tYXRpb246eWV0KTVUajgmSjkhQUFa").ConfigureAwait(false);
                        var incidentEntity = await this.incidentStorageProvider.GetAsync(activityToUpdate.IncidentNumber).ConfigureAwait(false);
                        var connector = new ConnectorClient(new Uri(incidentEntity.ServiceUrl), this.microsoftAppCredentials);
                        var activity = new Activity(ActivityTypes.Message)
                        {
                            Id = incidentEntity.TeamActivityId,
                            Conversation = new ConversationAccount { Id = incidentEntity.TeamConversationId },
                            Text = $"Incident: {activityToUpdate.IncidentNumber} current activity updated with: {activityToUpdate.Activity} by {turnContext.Activity.From.Name}",
                        };

                        await connector.Conversations.SendToConversationAsync(activity, cancellationToken);

                        // await turnContext.SendActivityAsync($"Incident: {activityToUpdate.IncidentNumber} current activity updated with: {activityToUpdate.Activity} by {turnContext.Activity.From.Name}");
                    }
                    else
                    {
                        var values = JsonConvert.DeserializeObject<ChangeTicketStatusPayload>(turnContext.Activity.Value.ToString());
                        var incidentDetails = await this.incidentStorageProvider.GetAsync(values.IncidentNumber, values.IncidentId).ConfigureAwait(false);
                        var updateObject = new Incident
                        {
                            Id = values.IncidentId,
                            Status = values.Action,
                        };
                        updateObject = await this.serviceNowProvider.UpdateIncidentAsync(updateObject, "U1ZDX3RlYW1zX2F1dG9tYXRpb246eWV0KTVUajgmSjkhQUFa").ConfigureAwait(false);
                        await this.incidentStorageProvider.AddAsync(new IncidentEntity { PartitionKey = values.IncidentNumber, RowKey = values.IncidentId, Status=values.Action}).ConfigureAwait(false);
                        updateObject.BridgeDetails = await this.conferenceBridgesStorageProvider.GetAsync(incidentDetails.BridgeId).ConfigureAwait(false);
                        var incidentCard = new IncidentCard(updateObject);
                        var attachment = incidentCard.GetIncidentAttachment();
                        if (values.Title != "Incident New")
                        {
                            attachment = incidentCard.GetIncidentAttachment(null, "Incident close", true);
                        }

                        if (values.Action != "1")
                        {
                            //    updateObject.BridgeDetails.Available = true;
                            //    await this.conferenceBridgesStorageProvider.AddAsync(updateObject.BridgeDetails).ConfigureAwait(false);
                        }

                        var connector = new ConnectorClient(new Uri(incidentDetails.ServiceUrl), this.microsoftAppCredentials);
                        var updateResponse = connector.Conversations.UpdateActivityAsync(incidentDetails.PersonalConversationId, incidentDetails.PersonalActivityId, (Activity)MessageFactory.Attachment(attachment), cancellationToken);
                        var updateResponseTeam = connector.Conversations.UpdateActivityAsync(incidentDetails.TeamConversationId, incidentDetails.TeamActivityId, (Activity)MessageFactory.Attachment(attachment), cancellationToken);
                        await Task.WhenAll(updateResponse, updateResponseTeam).ConfigureAwait(false);
                    }

                    break;
                case "HI":
                    await turnContext.SendActivityAsync(activity: MessageFactory.Attachment(WelcomeCard.GetWelcomeCardAttachment()), cancellationToken).ConfigureAwait(false);
                    break;

                default:
                    await this.dialog.RunAsync(turnContext, this.conversationState.CreateProperty<DialogState>(nameof(DialogState)), cancellationToken).ConfigureAwait(false);
                    break;
            }
        }

        /// <summary>
        /// Get the value of the searchText parameter in the messaging extension query.
        /// </summary>
        /// <param name="query">Contains messaging extension query keywords.</param>
        /// <returns>A value of the searchText parameter.</returns>
        private string GetSearchQueryString(MessagingExtensionQuery query)
        {
            var messageExtensionInputText = query.Parameters.FirstOrDefault(parameter => parameter.Name.Equals(SearchTextParameterName, StringComparison.OrdinalIgnoreCase));
            return messageExtensionInputText?.Value?.ToString();
        }

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
