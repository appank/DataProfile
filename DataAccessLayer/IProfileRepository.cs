using DataProfile.Models;
using System.Collections.Generic;

namespace DataProfile.DataAccessLayer
{
    public interface IProfileRepository
    {
        List<Profile> GetAllProfiles(); 
    }
}