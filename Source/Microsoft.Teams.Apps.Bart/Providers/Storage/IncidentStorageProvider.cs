// <copyright file="UserConfigurationStorageProvider.cs" company="Microsoft Corporation">
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
    public class IncidentStorageProvider : IIncidentStorageProvider
    {
        /// <summary>
        /// Table name in Azure table storage.
        /// </summary>
        private const string TableName = "Incidents";

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
        /// Initializes a new instance of the <see cref="IncidentStorageProvider"/> class.
        /// </summary>
        /// <param name="storageConnectionString">Azure Table Storage connection string.</param>
        /// <param name="telemetryClient">Telemetry client for logging events and errors.</param>
        public IncidentStorageProvider(string storageConnectionString, TelemetryClient telemetryClient)
        {
            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync(storageConnectionString));
            this.telemetryClient = telemetryClient;
        }

        /// <summary>
        /// Get user configuration.
        /// </summary>
        /// <param name="userObjectIdentifer">Active Directory object Id of user.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task<IncidentEntity> GetAsync(string partitionKey, string rowKey)
        {
            try
            {
                await this.EnsureInitializedAsync().ConfigureAwait(false);
                var retrieveOperation = TableOperation.Retrieve<IncidentEntity>(partitionKey, rowKey);
                var result = await this.cloudTable.ExecuteAsync(retrieveOperation).ConfigureAwait(false);
                return (IncidentEntity)result?.Result;
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }

        /// <summary>
        /// Add or update user configuration.
        /// </summary>
        /// <param name="incident">User configuration entity.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task<bool> AddAsync(IncidentEntity incident)
        {
            try
            {
                await this.EnsureInitializedAsync().ConfigureAwait(false);
                TableOperation insertOrMergeOperation = TableOperation.InsertOrMerge(incident);
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
        /// Get incidents.
        /// </summary>
        /// <param name="botCommand">Condition for searching.</param>
        /// <returns>A task that represents the work queued to execute.</returns>
        public async Task<List<IncidentEntity>> GetIncidentsAsync(string botCommand)
        {
            try
            {
                string condition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.NotEqual, string.Empty);
                switch (botCommand)
                {
                    case "newincidents":
                        condition = TableQuery.GenerateFilterCondition("Status", QueryComparisons.NotEqual, "1");
                        break;
                    case "suspendedincidents":
                        condition = TableQuery.GenerateFilterCondition("Status", QueryComparisons.NotEqual, "2");
                        break;
                    case "servicerestoredincidents":
                        condition = TableQuery.GenerateFilterCondition("Status", QueryComparisons.NotEqual, "3");
                        break;
                    case "allincidents":
                        condition = TableQuery.GenerateFilterCondition("Status", QueryComparisons.NotEqual, "1");
                        break;
                }
                await this.EnsureInitializedAsync().ConfigureAwait(false);
                var query = new TableQuery<IncidentEntity>().Where(condition);
                TableContinuationToken continuationToken = null;
                var incidents = new List<IncidentEntity>();

                do
                {
                    var queryResult = await this.cloudTable.ExecuteQuerySegmentedAsync(query, continuationToken).ConfigureAwait(false);
                    incidents.AddRange(queryResult?.Results);
                    continuationToken = queryResult?.ContinuationToken;
                }
                while (continuationToken != null);
                return incidents;
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
