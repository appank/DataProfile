using DataProfile.Models;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace DataProfile.DataAccessLayer
{

    public class ProfileRepository : IProfileRepository
    {
        private readonly string _connectionString;

       public ProfileRepository(IConfiguration configuration)
        {
            string? connStringFromConfig = configuration.GetConnectionString("DefaultConnection");

            if (string.IsNullOrEmpty(connStringFromConfig))
            {
                throw new InvalidOperationException("Connection string 'DefaultConnection' not found or is empty in appsettings.json.");
            }
            _connectionString = connStringFromConfig!; 
        }

        public List<Profile> GetAllProfiles()
        {
            var profiles = new List<Profile>();
            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    using (SqlCommand command = new SqlCommand("GetProfilesWithCTE", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        connection.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                profiles.Add(new Profile
                                {
                                    Id = Convert.ToInt32(reader["Id"]),
                                    FullName = reader["FullName"]?.ToString() ?? string.Empty,
                                    Email = reader["Email"]?.ToString() ?? string.Empty,
                                    PhoneNumber = reader["PhoneNumber"]?.ToString() ?? string.Empty,
                                    BirthDate = Convert.ToDateTime(reader["BirthDate"]),
                                    Address = reader["Address"]?.ToString() ?? string.Empty
                                });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error (gunakan ILogger jika sudah di-inject)
                Console.WriteLine($"Error in GetAllProfiles: {ex.Message}");
                // throw; // Atau tangani error sesuai kebutuhan
            }
            return profiles;
        }
    }
}