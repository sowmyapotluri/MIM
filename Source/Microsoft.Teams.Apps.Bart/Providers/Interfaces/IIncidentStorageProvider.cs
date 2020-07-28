// <copyright file="IIncidentStorageProvider.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Providers.Interfaces
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.Bart.Models.TableEntities;

    /// <summary>
    /// Storage provider for fetch, insert and update operation on Incidents table.
    /// </summary>
    public interface IIncidentStorageProvider
    {
        /// <summary>
        /// Add or update incidents.
        /// </summary>
        /// <param name="incidentEntity">Incident entity.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task<bool> AddAsync(IncidentEntity incidentEntity);

        /// <summary>
        /// Get incident.
        /// </summary>
        /// <param name="incidentNumber">Incident number.</param>
        /// <param name="incidentId">Incident id.</param>
        /// <returns>A task that represents the corresponding incident entity.</returns>
        Task<IncidentEntity> GetAsync(string incidentNumber, string incidentId);

        /// <summary>
        /// Get incidents.
        /// </summary>
        /// <param name="condition">Condition for searching.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task<List<IncidentEntity>> GetIncidentsAsync(string condition);

        /// <summary>
        /// Get incident based on incident number.
        /// </summary>
        /// <param name="incidentNumber">Incident number.</param>
        /// <returns>A task that represents corresponding incident.</returns>
        Task<IncidentEntity> GetAsync(string incidentNumber);
    }
}