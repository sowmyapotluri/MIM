// <copyright file="BotCommands.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart
{
    /// <summary>
    /// Bot commands.
    /// </summary>
    public static class BotCommands
    {
        /// <summary>
        /// Create incident command.
        /// </summary>
        public const string CreateIncident = "CREATE INCIDENT";

        /// <summary>
        /// Login command.
        /// </summary>
        public const string Login = "SIGN IN";

        /// <summary>
        /// Logout command.
        /// </summary>
        public const string Logout = "SIGN OUT";

        /// <summary>
        /// Help command.
        /// </summary>
        public const string Help = "HELP";

        /// <summary>
        /// Get workstreams command.
        /// </summary>
        public const string EditWorkstream = "EDIT WORKSTREAM";

        /// <summary>
        /// View workstreams command.
        /// </summary>
        public const string ViewWorkstream = "VIEW WORKSTREAM";

        /// <summary>
        /// Take a tour command.
        /// </summary>
        public const string TakeTour = "TAKE A TOUR";

        /// <summary>
        /// Action performed from card command.
        /// </summary>
        public const string CardAction = "CARDACTION";
    }
}
