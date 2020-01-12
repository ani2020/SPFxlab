// --------------------------------------------------------------------------------------------------------------------
// <copyright file="OppInfo.cs" company="Microsoft">
//   Copyright © 2020 Microsoft.
// </copyright>
// <summary>
//   Simple API to use in the SPFx Lab
// </summary>
// --------------------------------------------------------------------------------------------------------------------

using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Table;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace ReadyFeb2020.APIForSPFx
{
    public static class OppInfo
    {
        private const string TableName = "OpportunityInfo";

        /// <summary>
        /// Dummy method to test API deployment
        /// </summary>
        /// <param name="req"></param>
        /// <param name="log"></param>
        /// <returns></returns>
        [FunctionName("OppInfo")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null )] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            var SiteInfo = new {OpportunityId = "7-XDFSDFS3", 
                            Name = "My Project", 
                            Domain = "Apps", 
                            PrimaryProduct = "Azure",
                            Industry = "Banking"};
            var message = JsonConvert.SerializeObject(SiteInfo);

            return (ActionResult)new OkObjectResult(message);
        }

        /// <summary>
        /// Get Opportunity Info by Id
        /// </summary>
        /// <param name="req"></param>
        /// <param name="opp">Opportunity Info</param>
        /// <param name="log"></param>
        /// <param name="id">Id (Rowkey)</param>
        /// <returns></returns>
        [FunctionName("GetOppInfoById")]
        public static IActionResult GetOppInfoById(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "OppInfo/{id}")]HttpRequest req,
            [Table(TableName, "OpportunityInfo", "{id}", Connection = "AzureWebJobsStorage")] OpportunityInfoEntity opp,
            ILogger log, string id)
        {
            log.LogInformation("Getting opp item by id");
            if (opp == null)
            {
                log.LogInformation($"Item {id} not found");
                return new NotFoundResult();
            }
            return new OkObjectResult(Mappings.ToOpportunityInfo(opp));
        }

        /// <summary>
        /// Update Opportunity Info
        /// </summary>
        /// <param name="req"></param>
        /// <param name="OpportunityTable"></param>
        /// <param name="log"></param>
        /// <param name="id">id (rowkey) to update</param>
        /// <returns></returns>
        [FunctionName("UpdateOppInfoById")]
        public static async Task<IActionResult> UpdateOppInfoById(
            [HttpTrigger(AuthorizationLevel.Anonymous, "put", Route = "OppInfo/{id}")]HttpRequest req,
            [Table(TableName, Connection = "AzureWebJobsStorage")] CloudTable OpportunityTable,
            ILogger log, string id)
        {

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            var updated = JsonConvert.DeserializeObject<OpportunityInfoEntity>(requestBody);
            var findOperation = TableOperation.Retrieve<OpportunityInfoEntity>("OpportunityInfo", id);
            var findResult = await OpportunityTable.ExecuteAsync(findOperation);
            if (findResult.Result == null)
            {
                return new NotFoundResult();
            }
            var existingRow = (OpportunityInfoEntity)findResult.Result;

            if (!string.IsNullOrEmpty(updated.Industry))
            {
                existingRow.Industry = updated.Industry;
            }

            var replaceOperation = TableOperation.Replace(existingRow);
            await OpportunityTable.ExecuteAsync(replaceOperation);

            return new OkObjectResult(Mappings.ToOpportunityInfo(existingRow));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="req"></param>
        /// <param name="OpportunityInfoTable"></param>
        /// <param name="log"></param>
        /// <returns></returns>
        [FunctionName("CreateOpportunityInfo")]
        public static async Task<IActionResult>CreateCreateOpportunityInfo(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "OppInfo")]HttpRequest req, 
            [Table(TableName, Connection="AzureWebJobsStorage")] IAsyncCollector<OpportunityInfoEntity> OpportunityInfoTable,
            ILogger log)
        {
            log.LogInformation("Creating a new Opportunity list item");
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            var input = JsonConvert.DeserializeObject<OpportunityInfoEntity>(requestBody);

            var opp = new OpportunityInfo() 
                            { 
                                OpportunityId = input.OpportunityId,
                                Domain = input.Domain,
                                Name = input.Name, 
                                PrimaryProduct = input.PrimaryProduct,
                                Industry = input.Industry,
                            };
            await OpportunityInfoTable.AddAsync(Mappings.ToTableEntity(opp));
            return new OkObjectResult(opp);
        }    
        
    }

    /// <summary>
    /// Entity class for Table store
    /// </summary>
    public class OpportunityInfoEntity : TableEntity
    {
        public string OpportunityId { get; set; }
        public string Domain { get; set; }
        public string Name { get; set; }
        public string PrimaryProduct { get; set; }
        public string Industry { get; set; }
    }

    /// <summary>
    /// DTO
    /// </summary>
    public class OpportunityInfo
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("n");
        public string OpportunityId { get; set; }
        public string Domain { get; set; }
        public string Name { get; set; }
        public string PrimaryProduct { get; set; }
        public string Industry { get; set; }
    }


    public class OpportunityInfoModel
    {
        public string Id { get; set; }
        public string OpportunityId { get; set; }
        public string Domain { get; set; }
        public string Name { get; set; }
        public string PrimaryProduct { get; set; }
        public string Industry { get; set; }
    }

    /// <summary>
    /// DTO mapping
    /// </summary>
    public static class Mappings
    {
        /// <summary>
        /// Map DTO fields to Table fields
        /// </summary>
        /// <param name="opp">Opportunity Info DTO</param>
        /// <returns></returns>
        public static OpportunityInfoEntity ToTableEntity(OpportunityInfo opp)
        {
            return new OpportunityInfoEntity()
            {
                PartitionKey = "OpportunityInfo",
                RowKey = opp.Id,
                OpportunityId = opp.OpportunityId,
                Domain = opp.Domain,
                Name = opp.Name,
                PrimaryProduct = opp.PrimaryProduct,
                Industry = opp.Industry
            };
        }

        /// <summary>
        /// Map Table fields to DTO fields
        /// </summary>
        /// <param name="opp">Opportunity info Table entity</param>
        /// <returns></returns>
        public static OpportunityInfo ToOpportunityInfo(OpportunityInfoEntity opp)
        {
            return new OpportunityInfo()
            {
                Id = opp.RowKey,
                OpportunityId = opp.OpportunityId,
                Domain = opp.Domain,
                Name = opp.Name,
                PrimaryProduct = opp.PrimaryProduct,
                Industry = opp.Industry
            };
        }
    }
}
