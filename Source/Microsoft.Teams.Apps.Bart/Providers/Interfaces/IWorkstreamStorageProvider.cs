// <copyright file="IWorkstreamStorageProvider.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Providers.Interfaces
{
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.Bart.Models.TableEntities;

    /// <summary>
    /// Storage provider for fetch, insert and update operation on UserConfiguration table.
    /// </summary>
    public interface IWorkstreamStorageProvider
    {
        /// <summary>
        /// Add or update user configuration.
        /// </summary>
        /// <param name="workstreamEntity">User configuration entity.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task<bool> AddAsync(WorkstreamEntity workstreamEntity);

        /// <summary>
        /// Get user configuration.
        /// </summary>
        /// <param name="incidentNumber">Active Directory object Id of user.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task<WorkstreamEntity> GetAsync(string incidentNumber);
    }
}