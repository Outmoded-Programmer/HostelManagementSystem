using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace HostelManagementSystem
{
    public class HostelManager : IRoomManager
    {
        private readonly ExcelHelper _excelHelper;
        private const int TotalRooms = 8;
        private const int CapacityPerRoom = 4;
        private const string RoomsSheet = "Rooms";

        public HostelManager(ExcelHelper excelHelper)
        {
            _excelHelper = excelHelper ?? throw new ArgumentNullException(nameof(excelHelper));
        }

        public void InitializeRooms()
        {
            if (!_excelHelper.SheetExists(RoomsSheet))
            {
                DataTable table = new DataTable(RoomsSheet);
                table.Columns.Add("RoomNumber", typeof(int));
                table.Columns.Add("Occupied", typeof(int));

                for (int i = 1; i <= TotalRooms; i++)
                {
                    DataRow row = table.NewRow();
                    row["RoomNumber"] = i;
                    row["Occupied"] = 0;
                    table.Rows.Add(row);
                }

                _excelHelper.CreateExcelFile(new[] { table });
            }
        }

        public bool RoomExists(int roomNumber)
        {
            try
            {
                DataTable roomsTable = _excelHelper.ReadData(RoomsSheet);
                return roomsTable.AsEnumerable()
                    .Any(row => Convert.ToInt32(row["RoomNumber"]) == roomNumber);
            }
            catch
            {
                return false;
            }
        }

        public bool IsRoomFull(int roomNumber)
        {
            try
            {
                DataTable roomsTable = _excelHelper.ReadData(RoomsSheet);
                DataRow room = roomsTable.AsEnumerable()
                    .FirstOrDefault(row => Convert.ToInt32(row["RoomNumber"]) == roomNumber);

                return room != null && Convert.ToInt32(room["Occupied"]) >= CapacityPerRoom;
            }
            catch
            {
                return true;
            }
        }

        public bool AssignStudentToRoom(string studentId, int roomNumber)
        {
            if (roomNumber < 1 || roomNumber > TotalRooms || string.IsNullOrWhiteSpace(studentId))
                return false;

            try
            {
                DataTable roomsTable = _excelHelper.ReadData(RoomsSheet);
                DataRow room = roomsTable.AsEnumerable()
                    .FirstOrDefault(row => Convert.ToInt32(row["RoomNumber"]) == roomNumber);

                if (room == null)
                    return false;

                int occupied = Convert.ToInt32(room["Occupied"]);
                if (occupied >= CapacityPerRoom)
                    return false;

                room["Occupied"] = occupied + 1;
                _excelHelper.WriteData(RoomsSheet, roomsTable);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool RemoveStudentFromRoom(string studentId, int roomNumber)
        {
            if (roomNumber < 1 || roomNumber > TotalRooms || string.IsNullOrWhiteSpace(studentId))
                return false;

            try
            {
                DataTable roomsTable = _excelHelper.ReadData(RoomsSheet);
                DataRow room = roomsTable.AsEnumerable()
                    .FirstOrDefault(row => Convert.ToInt32(row["RoomNumber"]) == roomNumber);

                if (room == null)
                    return false;

                int occupied = Convert.ToInt32(room["Occupied"]);
                if (occupied <= 0)
                    return false;

                room["Occupied"] = occupied - 1;
                _excelHelper.WriteData(RoomsSheet, roomsTable);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void DisplayRoomStatus()
        {
            try
            {
                DataTable roomsTable = _excelHelper.ReadData(RoomsSheet);

                Console.WriteLine("\nRoom Status:");
                Console.WriteLine("Room\tOccupied\tAvailable");
                foreach (DataRow room in roomsTable.Rows)
                {
                    int roomNum = Convert.ToInt32(room["RoomNumber"]);
                    int occupied = Convert.ToInt32(room["Occupied"]);
                    Console.WriteLine($"{roomNum}\t{occupied}\t\t{CapacityPerRoom - occupied}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error displaying room status: {ex.Message}");
            }
        }

        public List<int> GetAvailableRooms()
        {
            var availableRooms = new List<int>();
            try
            {
                DataTable roomsTable = _excelHelper.ReadData(RoomsSheet);
                foreach (DataRow room in roomsTable.Rows)
                {
                    int occupied = Convert.ToInt32(room["Occupied"]);
                    if (occupied < CapacityPerRoom)
                    {
                        availableRooms.Add(Convert.ToInt32(room["RoomNumber"]));
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting available rooms: {ex.Message}");
            }
            return availableRooms;
        }
    }
}