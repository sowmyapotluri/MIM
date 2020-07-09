// <copyright file="WorkstreamStorageProvider.cs" company="Microsoft Corporation">
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
    public class WorkstreamStorageProvider : IWorkstreamStorageProvider
    {
        /// <summary>
        /// Table name in Azure table storage.
        /// </summary>
        private const string TableName = "Workstreams";

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
        /// Initializes a new instance of the <see cref="WorkstreamStorageProvider"/> class.
        /// </summary>
        /// <param name="storageConnectionString">Azure Table Storage connection string.</param>
        /// <param name="telemetryClient">Telemetry client for logging events and errors.</param>
        public WorkstreamStorageProvider(string storageConnectionString, TelemetryClient telemetryClient)
        {
            this.initializeTask = new Lazy<Task>(() => this.InitializeAsync(storageConnectionString));
            this.telemetryClient = telemetryClient;
        }

        /// <summary>
        /// Get a workstream.
        /// </summary>
        /// <param name="incidentNumber">Incident number.</param>
        /// <param name="id">Workstream id.</param>
        /// <returns>A workstream item.</returns>
        public async Task<WorkstreamEntity> GetAsync(string incidentNumber, string id)
        {
            try
            {
                await this.EnsureInitializedAsync().ConfigureAwait(false);
                var retrieveOperation = TableOperation.Retrieve<IncidentEntity>(incidentNumber, id);
                var result = await this.cloudTable.ExecuteAsync(retrieveOperation).ConfigureAwait(false);
                return (WorkstreamEntity)result?.Result;
            }
            catch (Exception ex)
            {
                this.telemetryClient.TrackException(ex);
                throw;
            }
        }

        /// <summary>
        /// Get all workstream.
        /// </summary>
        /// <param name="incidentNumber">Incident number.</param>
        /// <returns>List of all workstreams.</returns>
        public async Task<List<WorkstreamEntity>> GetAllAsync(string incidentNumber)
        {
            try
            {
                await this.EnsureInitializedAsync().ConfigureAwait(false);
                string partitionKeyCondition = TableQuery.GenerateFilterCondition("PartitionKey", QueryComparisons.Equal, incidentNumber);
                var query = new TableQuery<WorkstreamEntity>().Where(partitionKeyCondition);
                TableContinuationToken continuationToken = null;
                var queryResult = await this.cloudTable.ExecuteQuerySegmentedAsync(query, continuationToken).ConfigureAwait(false);
                return queryResult.Results;
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
        /// <param name="workstream">User configuration entity.</param>
        /// <returns>A task that represents whether it was successfull or not.</returns>
        public async Task<bool> AddAsync(WorkstreamEntity workstream)
        {
            try
            {
                await this.EnsureInitializedAsync().ConfigureAwait(false);
                TableOperation insertOrMergeOperation = TableOperation.InsertOrReplace(workstream);
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
        /// Delete a workstream.
        /// </summary>
        /// <param name="workstream">User configuration entity.</param>
        /// <returns>Boolean value to confirm deletion.</returns>
        public async Task<bool> DeleteAsync(WorkstreamEntity workstream)
        {
            try
            {
                await this.EnsureInitializedAsync().ConfigureAwait(false);
                TableOperation deleteOperation = TableOperation.Delete(workstream);
                TableResult result = await this.cloudTable.ExecuteAsync(deleteOperation).ConfigureAwait(false);
                return result.Result != null;
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
