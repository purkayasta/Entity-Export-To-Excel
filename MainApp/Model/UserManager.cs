using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MainApp.Model
{
    public class UserManager
    {
        public string Id { get; set; }
        public string Username { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Password { get; set; }
        public DateTime CreatedOn { get; set; }

        public static List<UserManager> SeedData()
        {
            return new List<UserManager>
            {
                new UserManager
                {
                    Id = Guid.NewGuid().ToString(),
                    FirstName = "Test 1",
                    LastName = "Test 1",
                    Username = "TestUsername1",
                    Password = "123",
                    CreatedOn = DateTime.UtcNow
                },
                new UserManager
                {
                    Id = Guid.NewGuid().ToString(),
                    FirstName = "Test 2",
                    LastName = "Test 2",
                    Username = "TestUsername2",
                    Password = "123",
                    CreatedOn = DateTime.UtcNow
                },
                new UserManager
                {
                    Id = Guid.NewGuid().ToString(),
                    FirstName = "Test 3",
                    LastName = "Test 3",
                    Username = "TestUsername3",
                    Password = "123",
                    CreatedOn = DateTime.UtcNow
                },new UserManager
                {
                    Id = Guid.NewGuid().ToString(),
                    FirstName = "Test 4",
                    LastName = "Test 4",
                    Username = "TestUsername4",
                    Password = "123",
                    CreatedOn = DateTime.UtcNow
                }
            };
        }
    }
}
