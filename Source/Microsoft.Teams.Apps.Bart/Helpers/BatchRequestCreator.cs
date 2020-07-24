// <copyright file="BatchRequestCreator.cs" company="Microsoft Corporation">
// Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.Bart.Helpers
{
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Teams.Apps.Bart.Models;

    /// <summary>
    /// creating <see cref="BatchRequestCreator"/> class.
    /// </summary>
    public class BatchRequestCreator
    {
        /// <summary>
        /// Gets or sets this metod to get unique id of each request in batch call.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets this metod to get method of each request in batch call.
        /// </summary>
        public string Method { get; set; }

        /// <summary>
        /// Gets or sets this metod to get url of each request in batch call.
        /// </summary>
        public string URL { get; set; }

        /// <summary>
        /// This method is to create batch request for Graph Api calls.
        /// </summary>
        /// <param name="incidents">List of incident.</param>
        /// <returns>A <see cref="dynamic"/> representing the request payload for batch api call.</returns>
        public dynamic CreateBatchRequestPayloadForDetails(List<IncidentListObject> incidents)
        {
           List<dynamic> request = new List<dynamic>();
           List<User> users = new List<User>();
           foreach (IncidentListObject incidentObject in incidents)
            {
                if (users.Select((user) => user.Id == incidentObject.RequestedBy.Id).Count() == 0)
                {
                    users.Add(incidentObject.RequestedBy);
                    BatchRequestCreator batchRequestCreator = new BatchRequestCreator()
                    {
                        Id = incidentObject.RequestedBy.Id,
                        Method = "GET",
                        URL = "/users/" + incidentObject.RequestedBy.Id + "/photo/$value",
                    };
                    request.Add(batchRequestCreator);
                }

                if (users.Select((user) => user.Id == incidentObject.AssignedTo.Id).Count() == 0)
                {
                    users.Add(incidentObject.AssignedTo);
                    BatchRequestCreator batchRequestCreator = new BatchRequestCreator()
                    {
                        Id = incidentObject.AssignedTo.Id,
                        Method = "GET",
                        URL = "/users/" + incidentObject.AssignedTo.Id + "/photo/$value",
                    };
                    request.Add(batchRequestCreator);
                }
            }

           return request;
        }
    }
}
