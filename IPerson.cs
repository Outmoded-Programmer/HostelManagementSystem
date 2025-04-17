using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HostelManagementSystem
{
    internal interface IPerson
    {
        string ID { get; }
        string Name { get; }
        string Email { get; }
        void ViewDetails();
    }
}
