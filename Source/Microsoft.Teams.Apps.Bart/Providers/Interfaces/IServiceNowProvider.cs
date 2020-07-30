// <copyright file="IServiceNowProvider.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Providers.Interfaces
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.Bart.Models;

    /// <summary>
    /// Provider which exposes methods required for incident creation.
    /// </summary>
    public interface IServiceNowProvider
    {
        /// <summary>
        /// Create new incident.
        /// </summary>
        /// <param name="incident"><see cref="Incident"/> object. </param>
        /// <returns>Event response object.</returns>
        Task<dynamic> CreateIncidentAsync(Incident incident);

        /// <summary>
        /// Update new incident.
        /// </summary>
        /// <param name="incident"><see cref="Incident"/> object. </param>
        /// <returns>Event response object.</returns>
        Task<dynamic> UpdateIncidentAsync(Incident incident);

        /// <summary>
        /// Get incidents.
        /// </summary>
        /// <param name="currentMonthStartDate">Month start date. </param>
        /// <param name="currentMonthEndDate">Month end date. </param>
        /// <returns>Event response object.</returns>
        Task<dynamic> SearchIncidentAsync(DateTime currentMonthStartDate, DateTime currentMonthEndDate);

        /// <summary>
        /// Get incidents based on factors.
        /// </summary>
        /// <param name="commandId">Bot command for searching. </param>
        /// <param name="searchQuery">Query for searching. </param>
        /// <returns>Event response object.</returns>
        Task<dynamic> GetIncidentAsync(string commandId, string searchQuery);

        /// <summary>
        /// Get incident.
        /// </summary>
        /// <param name="incidentId">Incident id. </param>
        /// <returns>Event response object.</returns>
        Task<dynamic> GetIncidentAsync(string incidentId);
    }
}
