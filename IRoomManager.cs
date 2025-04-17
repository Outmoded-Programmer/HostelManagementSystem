using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HostelManagementSystem
{
    public interface IRoomManager
    {
        void InitializeRooms();
        bool RoomExists(int roomNumber);
        bool IsRoomFull(int roomNumber);
        bool AssignStudentToRoom(string studentId, int roomNumber);
        bool RemoveStudentFromRoom(string studentId, int roomNumber);
        void DisplayRoomStatus();
    }
}
