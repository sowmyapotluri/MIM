// <copyright file="StatusStorageProvider.cs" company="Microsoft Corporation">
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
    public class StatusStorageProvider : IStatusStorageProvider
    {
        /// <summary>
        /// Table name in Azure table storage.
        /// </summary>
        private const string TableName = "StatusConfiguration";

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
        /// Initializes a new instance of the <see cref="StatusStorageProvider"/> class.
        /// </summary>
        /// <param name="storageConnectionString">Azure Table Storage connection string.</param>
        /// <param name="telemetryClient">Telemetry client for logging events and errors.</param>
        public StatusStorageProvider(string storageConnectionString, TelemetryClient telemetryClient)
        {
            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync(storageConnectionString));
            this.telemetryClient = telemetryClient;
        }

        /// <summary>
        /// Add or update status.
        /// </summary>
        /// <param name="statusEntity">Status entity.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task<bool> AddAsync(StatusEntity statusEntity)
        {
            try
            {
                await this.EnsureInitializedAsync().ConfigureAwait(false);
                TableOperation insertOrMergeOperation = TableOperation.InsertOrReplace(statusEntity);
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
        /// Get statuses.
        /// </summary>
        /// <returns>A task that represents the list of statuses.</returns>
        public async Task<List<StatusEntity>> GetAsync()
        {
            try
            {
                await this.EnsureInitializedAsync().ConfigureAwait(false);
                TableContinuationToken token = null;
                var statuses = await this.cloudTable.ExecuteQuerySegmentedAsync(new TableQuery<StatusEntity>(), token).ConfigureAwait(false);
                return statuses?.Results;
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


