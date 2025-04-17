using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HostelManagementSystem
{
    public class Student : ComplaintHandler
    {
        public int RoomNumber { get; private set; }
        public bool FeesPaid { get; private set; }

        public Student(string id, string name, string email, string password, ExcelHelper excelHelper)
            : base(id, name, email, password, excelHelper)
        {
            DataRow studentData = FindPersonInSheet("Students", id);
            if (studentData != null)
            {
                RoomNumber = Convert.ToInt32(studentData["RoomNumber"]);

                // Fix: Safely convert "0"/"1"/"true"/"false" to boolean
                string feesPaidValue = studentData["FeesPaid"].ToString().ToLower();
                FeesPaid = feesPaidValue == "1" || feesPaidValue == "true";
            }
        }

        public void PayFees(decimal amount)
        {
            DataTable feesTable = ExcelHelper.ReadData("Fees");
            DataRow newFee = feesTable.NewRow();
            newFee["StudentID"] = this.ID;
            newFee["Amount"] = amount;
            newFee["PaymentDate"] = DateTime.Now.ToString("yyyy-MM-dd");
            newFee["DueDate"] = DateTime.Now.AddMonths(1).ToString("yyyy-MM-dd");

            ExcelHelper.AppendData("Fees", newFee);

            ExcelHelper.UpdateRow("Students", "ID", this.ID,
                new Dictionary<string, string> { { "FeesPaid", "True" } });

            this.FeesPaid = true;
            Console.WriteLine($"Payment of {amount} received. Thank you!");
        }

        public override void ViewDetails()
        {
            Console.WriteLine($"ID: {ID}");
            Console.WriteLine($"Name: {Name}");
            Console.WriteLine($"Email: {Email}");
            Console.WriteLine($"Room Number: {RoomNumber}");
            Console.WriteLine($"Fees Paid: {(FeesPaid ? "Yes" : "No")}");
        }
    }
}
