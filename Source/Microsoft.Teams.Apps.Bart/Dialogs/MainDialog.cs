// <copyright file="MainDialog.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Dialogs
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using AdaptiveCards;
    using Microsoft.ApplicationInsights;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Builder.Dialogs;
    using Microsoft.Bot.Schema;
    using Microsoft.Bot.Schema.Teams;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Teams.Apps.Bart;
    using Microsoft.Teams.Apps.Bart.Cards;
    using Microsoft.Teams.Apps.Bart.Dialogs;
    using Microsoft.Teams.Apps.Bart.Helpers;
    using Microsoft.Teams.Apps.Bart.Models;
    using Microsoft.Teams.Apps.Bart.Providers.Interfaces;
    using Microsoft.Teams.Apps.Bart.Providers.Storage;
    using Microsoft.Teams.Apps.Bart.Resources;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// Acts as root dialog for processing commands received from user.
    /// </summary>
    public class MainDialog : LogoutDialog
    {
        /// <summary>
        /// Helper which exposes methods required for meeting creation process.
        /// </summary>
        private readonly IServiceNowProvider serviceNowProvider;

        /// <summary>
        /// Storage provider to perform fetch, insert and update operation on ActivityEntities table.
        /// </summary>
        private readonly IActivityStorageProvider activityStorageProvider;

        /// <summary>
        /// Helper for generating and validating JWT token.
        /// </summary>
        private readonly ITokenHelper tokenHelper;

        /// <summary>
        /// Telemetry client for logging events and errors.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Storage provider to perform fetch operation on UserConfiguration table.
        /// </summary>
        private readonly IUserConfigurationStorageProvider userConfigurationStorageProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="MainDialog"/> class.
        /// </summary>
        /// <param name="configuration">Application configuration.</param>
        /// <param name="meetingProvider">Helper which exposes methods required for meeting creation process.</param>
        /// <param name="activityStorageProvider">Storage provider to perform fetch, insert and update operation on ActivityEntities table.</param>
        /// <param name="favoriteStorageProvider">Storage provider to perform fetch, insert, update and delete operation on UserFavorites table.</param>
        /// <param name="tokenHelper">Helper for generating and validating JWT token.</param>
        /// <param name="telemetryClient">Telemetry client for logging events and errors.</param>
        /// <param name="userConfigurationStorageProvider">Storage provider to perform fetch operation on UserConfiguration table.</param>
        /// <param name="meetingHelper">Helper class which exposes methods required for meeting creation.</param>
        public MainDialog(IConfiguration configuration, IServiceNowProvider serviceNowProvider, IActivityStorageProvider activityStorageProvider, ITokenHelper tokenHelper, TelemetryClient telemetryClient, IUserConfigurationStorageProvider userConfigurationStorageProvider)
            : base(nameof(MainDialog), configuration["ConnectionName"], telemetryClient)
        {
            this.tokenHelper = tokenHelper;
            this.telemetryClient = telemetryClient;
            this.activityStorageProvider = activityStorageProvider;
            this.serviceNowProvider = serviceNowProvider;
            this.userConfigurationStorageProvider = userConfigurationStorageProvider;
            this.AddDialog(new OAuthPrompt(
                 nameof(OAuthPrompt),
                 new OAuthPromptSettings
                 {
                     ConnectionName = this.ConnectionName,
                     Text = Strings.SignInRequired,
                     Title = Strings.SignIn,
                     Timeout = 120000,
                 }));
            this.AddDialog(
                new WaterfallDialog(
                    nameof(WaterfallDialog),
                    new WaterfallStep[] { this.PromptStepAsync, this.CommandStepAsync, this.ProcessStepAsync }));
            this.InitialDialogId = nameof(WaterfallDialog);
        }

        /// <summary>
        /// Prompts sign in card.
        /// </summary>
        /// <param name="stepContext">Context object passed in to a WaterfallStep.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task<DialogTurnResult> PromptStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            stepContext.Values["command"] = stepContext.Context.Activity.Text?.Trim();
            if (stepContext.Context.Activity.Text == null && stepContext.Context.Activity.Value != null && stepContext.Context.Activity.Type == "message")
            {
                stepContext.Values["command"] = JToken.Parse(stepContext.Context.Activity.Value.ToString()).SelectToken("text").ToString().Trim();
            }

            return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// To get access token, calling prompt again.
        /// </summary>
        /// <param name="stepContext">Context object passed in to a WaterfallStep.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task<DialogTurnResult> CommandStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            var tokenResponse = (TokenResponse)stepContext.Result;
            if (!string.IsNullOrEmpty(tokenResponse?.Token))
            {
                if (stepContext.Values.ContainsKey("command"))
                {
                    stepContext.Context.Activity.Text = (string)stepContext.Values["command"] ?? string.Empty;
                }

                return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken).ConfigureAwait(false);
            }
            else
            {
                await stepContext.Context.SendActivityAsync(MessageFactory.Text(Strings.CantLogIn), cancellationToken).ConfigureAwait(false);
                return await stepContext.EndDialogAsync().ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Process the command user typed.
        /// </summary>
        /// <param name="stepContext">Context object passed in to a WaterfallStep.</param>
        /// <param name="cancellationToken">Propagates notification that operations should be canceled.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task<DialogTurnResult> ProcessStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            if (stepContext.Result != null)
            {
                var tokenResponse = stepContext.Result as TokenResponse;
                if (!string.IsNullOrEmpty(tokenResponse.Token))
                {
                    var command = (string)stepContext.Values["command"] ?? string.Empty;
                    switch (command.ToUpperInvariant())
                    {
                        case "HI":
                            await stepContext.Context.SendActivityAsync(activity: MessageFactory.Attachment(WelcomeCard.GetWelcomeCardAttachment()), cancellationToken).ConfigureAwait(false);
                            break;

                        case BotCommands.Help:
                            break;

                        case string message when message.Equals(BotCommands.Login, StringComparison.OrdinalIgnoreCase) || message.Equals(BotCommands.Logout, StringComparison.OrdinalIgnoreCase):
                            break;

                        default:
                            await stepContext.Context.SendActivityAsync(MessageFactory.Text(Strings.CommandNotRecognized.Replace("{command}", command, StringComparison.CurrentCulture)), cancellationToken).ConfigureAwait(false);
                            break;
                    }
                }
            }
            else
            {
                await stepContext.Context.SendActivityAsync(MessageFactory.Text(Strings.CantLogIn), cancellationToken).ConfigureAwait(false);
            }

            return await stepContext.EndDialogAsync().ConfigureAwait(false);
        }
    }
}
