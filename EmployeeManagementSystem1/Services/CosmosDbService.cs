using Microsoft.Azure.Cosmos;
using System.Threading.Tasks;
using System.Collections.Generic;
using EmployeeManagementSystem.Models;
using Microsoft.Extensions.Configuration;
using System.Collections.Concurrent;
using System.ComponentModel;
using Container = Microsoft.Azure.Cosmos.Container;

namespace EmployeeManagementSystem.Services
{
    public class CosmosDbService : ICosmosDbService
    {
        private Container _container;

        public CosmosDbService(CosmosClient cosmosClient, string databaseName, string containerName)
        {
            this._container = cosmosClient.GetContainer(databaseName, containerName);
        }

        public async Task AddEmployeeAsync(EmployeeBasicDetails employee)
        {
            await this._container.CreateItemAsync(employee, new PartitionKey(employee.EmployeeID));
        }

        public async Task DeleteEmployeeAsync(string id)
        {
            await this._container.DeleteItemAsync<EmployeeBasicDetails>(id, new PartitionKey(id));
        }

        public async Task<EmployeeBasicDetails> GetEmployeeAsync(string id)
        {
            try
            {
                ItemResponse<EmployeeBasicDetails> response = await this._container.ReadItemAsync<EmployeeBasicDetails>(id, new PartitionKey(id));
                return response.Resource;
            }
            catch (CosmosException ex) when (ex.StatusCode == System.Net.HttpStatusCode.NotFound)
            {
                return null;
            }
        }

        
           public async Task<IEnumerable<EmployeeBasicDetails>> GetEmployeesAsync(string queryString)
            {
                var query = this._container.GetItemQueryIterator<EmployeeBasicDetails>(new QueryDefinition(queryString));
                List<EmployeeBasicDetails> results = new List<EmployeeBasicDetails>();
                while (query.HasMoreResults)
                {
                    var response = await query.ReadNextAsync();
                    results.AddRange(response.ToList());
                }
                return results;
            }

            public async Task UpdateEmployeeAsync(string id, EmployeeBasicDetails employee)
            {
                await this._container.UpsertItemAsync(employee, new PartitionKey(id));
            }
        }
    }