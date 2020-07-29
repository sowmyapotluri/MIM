// <copyright file="ConferenceBridgesStorageProvider.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Providers.Storage
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.Teams.Apps.Bart.Models.TableEntities;
    using Microsoft.Teams.Apps.Bart.Providers.Interfaces;
    using Microsoft.WindowsAzure.Storage;
    using Microsoft.WindowsAzure.Storage.Table;

    /// <summary>
    /// Storage provider for fetch, insert and update operation on UserConfiguration table.
    /// </summary>
    public class ConferenceBridgesStorageProvider : IConferenceBridgesStorageProvider
    {
        /// <summary>
        /// Table name in Azure table storage.
        /// </summary>
        private const string TableName = "ConferenceRooms";

        /// <summary>
        /// Task for initialization.
        /// </summary>
        private readonly Lazy<Task> initializeTask;

        /// <summary>
        /// Telemetry client for logging events and errors.
        /// </summary>
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Provides a service client for accessing the Microsoft Azure Table service.
        /// </summary>
        private CloudTableClient cloudTableClient;

        /// <summary>
        /// Represents a table in the Microsoft Azure Table service.
        /// </summary>
        private CloudTable cloudTable;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConferenceBridgesStorageProvider"/> class.
        /// </summary>
        /// <param name="storageConnectionString">Azure Table Storage connection string.</param>
        /// <param name="telemetryClient">Telemetry client for logging events and errors.</param>
        public ConferenceBridgesStorageProvider(string storageConnectionString, TelemetryClient telemetryClient)
        {
            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync(storageConnectionString));
            this.telemetryClient = telemetryClient;
        }

        /// <summary>
        /// Add or update conference bridge details.
        /// </summary>
        /// <param name="conferenceRoom">Conference room entity.</param>
        /// <returns>A task that represents status of execution.</returns>
        public async Task<bool> AddAsync(ConferenceRoomEntity conferenceRoom)
        {
            try
            {
                await this.EnsureInitializedAsync().ConfigureAwait(false);
                TableOperation insertOrMergeOperation = TableOperation.InsertOrReplace(conferenceRoom);
                TableResult result = await this.cloudTable.ExecuteAsync(insertOrMergeOperation).ConfigureAwait(false);
                return result.Result != null;
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }

        /// <summary>
        /// Get all conference bridge data.
        /// </summary>
        /// <returns>A task that represents all conference bridge details.</returns>
        public async Task<List<ConferenceRoomEntity>> GetAsync()
        {
            try
            {
                await this.EnsureInitializedAsync().ConfigureAwait(false);
                TableContinuationToken token = null;
                string filter = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, "conferencerooms");
                TableQuery<ConferenceRoomEntity> tableQuery =
                       new TableQuery<ConferenceRoomEntity>().Where(filter); //.Select(new List<string> { });
                var room = await this.cloudTable.ExecuteQuerySegmentedAsync(tableQuery, token).ConfigureAwait(false);
                return room?.Results;
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }

        /// <summary>
        /// Get conference bridge data.
        /// </summary>
        /// <param name="bridge">Unique bridge code.</param>
        /// <returns>A task that represents corresponding conference bridge details.</returns>
        public async Task<ConferenceRoomEntity> GetAsync(string bridge)
        {
            try
            {
                await this.EnsureInitializedAsync().ConfigureAwait(false);
                var retrieveOperation = TableOperation.Retrieve<ConferenceRoomEntity>("conferencerooms", bridge);
                var room = await this.cloudTable.ExecuteAsync(retrieveOperation).ConfigureAwait(false);
                return (ConferenceRoomEntity)room?.Result;
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }

        /// <summary>
        /// Ensure table storage connection is initialized.
        /// </summary>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task EnsureInitializedAsync()
        {
            await this.initializeTask.Value;
        }

        /// <summary>
        /// Create tables if it doesn't exists.
        /// </summary>
        /// <param name="connectionString">Storage account connection string.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        private async Task InitializeAsync(string connectionString)
        {
            try
            {
                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
                this.cloudTableClient = storageAccount.CreateCloudTableClient();
                this.cloudTable = this.cloudTableClient.GetTableReference(TableName);
                await this.cloudTable.CreateIfNotExistsAsync().ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }
    }
}