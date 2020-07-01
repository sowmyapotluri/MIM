// <copyright file="IConferenceBridgesStorageProvider.cs" company="Microsoft Corporation">
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
    public interface IConferenceBridgesStorageProvider
    {
        /// <summary>
        /// Add or update user configuration.
        /// </summary>
        /// <param name="conferenceRoom">Conference room entity.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task<bool> AddAsync(ConferenceRoomEntity conferenceRoom);

        /// <summary>
        /// Get user configuration.
        /// </summary>
        /// <param name="bridge">Active Directory object Id of user.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task<ConferenceRoomEntity> GetAsync(string bridge);

        /// <summary>
        /// Get user configuration.
        /// </summary>
        /// <returns>A task that represents the work queued to execute.</returns>
        Task<List<ConferenceRoomEntity>> GetAsync();
    }
}