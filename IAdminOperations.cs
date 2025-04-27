using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HostelManagementSystem
{
    public interface IAdminOperations : IStaffOperations
    {
        void AddStudent(int id , string name, string email, string password);
        void ViewAllStudents();
        //void UpdateStudent(string studentId, string studentName, string studentEmail, string studentPassword);
        void UpdateStudent(string studentId , string studentName, string studentEmail , string studentPassword);
        void AssignRoom(string studentId, int roomNumber);
        void RemoveStudent(string studentId);
        void ViewRoomStatus();
        void ViewAllComplaints();
        void UpdateComplaintStatus(string complaintId, string newStatus);
        void AddStaffManager(string name, string email, string password);
    }
}
