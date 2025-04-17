using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HostelManagementSystem
{
    public interface IStaffOperations
    {
        void ViewAllStaff();
        void AddStaff(string name, string email, string password, string position);
        void RemoveStaff(string staffId);
    }
}
