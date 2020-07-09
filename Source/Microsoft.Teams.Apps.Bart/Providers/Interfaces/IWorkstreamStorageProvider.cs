// <copyright file="IWorkstreamStorageProvider.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Providers.Interfaces
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.Bart.Models.TableEntities;

    /// <summary>
    /// Storage provider for fetch, insert and update operation on UserConfiguration table.
    /// </summary>
    public interface IWorkstreamStorageProvider
    {
        /// <summary>
        /// Add or update workstream.
        /// </summary>
        /// <param name="workstreamEntity">User configuration entity.</param>
        /// <returns>A task that represents whether it was successfull or not.</returns>
        Task<bool> AddAsync(WorkstreamEntity workstreamEntity);

        /// <summary>
        /// Get all workstream.
        /// </summary>
        /// <param name="incidentNumber">Incident number.</param>
        /// <returns>List of all workstreams.</returns>

        Task<List<WorkstreamEntity>> GetAllAsync(string incidentNumber);

        /// <summary>
        /// Get a workstream.
        /// </summary>
        /// <param name="incidentNumber">Incident number.</param>
        /// <param name="id">Workstream id.</param>
        /// <returns>A workstream item.</returns>
        Task<WorkstreamEntity> GetAsync(string incidentNumber, string id);

        /// <summary>
        /// Delete a workstream.
        /// </summary>
        /// <param name="workstreamEntity">User configuration entity.</param>
        /// <returns>Boolean value to confirm deletion.</returns>
        Task<bool> DeleteAsync(WorkstreamEntity workstreamEntity);
    }
}