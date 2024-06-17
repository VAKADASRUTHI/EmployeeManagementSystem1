using System.Collections.Generic;
using System.Threading.Tasks;
using EmployeeManagementSystem.Models;

namespace EmployeeManagementSystem.Services
{
    public interface ICosmosDbService
    {
        Task<IEnumerable<EmployeeBasicDetails>> GetEmployeesAsync(string queryString);
        Task<EmployeeBasicDetails> GetEmployeeAsync(string id);
        Task AddEmployeeAsync(EmployeeBasicDetails employee);
        Task UpdateEmployeeAsync(string id, EmployeeBasicDetails employee);
        Task DeleteEmployeeAsync(string id);
    }
}