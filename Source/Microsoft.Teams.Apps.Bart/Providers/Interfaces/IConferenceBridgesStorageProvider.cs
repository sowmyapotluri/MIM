// <copyright file="IConferenceBridgesStorageProvider.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Providers.Interfaces
{
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.Bart.Models.TableEntities;

    /// <summary>
    /// Storage provider for fetch, insert and update operation on ConferenceBridge table.
    /// </summary>
    public interface IConferenceBridgesStorageProvider
    {
        /// <summary>
        /// Add or update conference bridge details.
        /// </summary>
        /// <param name="conferenceRoom">Conference room entity.</param>
        /// <returns>A task that represents status of execution.</returns>
        Task<bool> AddAsync(ConferenceRoomEntity conferenceRoom);

        /// <summary>
        /// Get conference bridge data.
        /// </summary>
        /// <param name="bridge">Unique bridge code.</param>
        /// <returns>A task that represents corresponding conference bridge details.</returns>
        Task<ConferenceRoomEntity> GetAsync(string bridge);

        /// <summary>
        /// Get all conference bridge data.
        /// </summary>
        /// <returns>A task that represents all conference bridge details.</returns>
        Task<List<ConferenceRoomEntity>> GetAsync();
    }
}