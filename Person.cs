using System.Data;

namespace HostelManagementSystem
{
    public abstract class Person : IPerson
    {
        public string ID { get; protected set; }
        public string Name { get; protected set; }
        public string Email { get; protected set; }
        protected string Password { get; set; }
        protected ExcelHelper ExcelHelper { get; }

        protected Person(string id, string name, string email, string password, ExcelHelper excelHelper)
        {
            ID = id;
            Name = name;
            Email = email;
            Password = password;
            ExcelHelper = excelHelper;
        }

        public virtual bool Login(string email, string password)
        {
            return Email == email && Password == password;
        }

        public abstract void ViewDetails();

        protected DataRow FindPersonInSheet(string sheetName, string id)
        {
            DataTable dt = ExcelHelper.ReadData(sheetName);
            foreach (DataRow row in dt.Rows)
            {
                if (row["ID"].ToString() == id)
                {
                    return row;
                }
            }
            return null;
        }
    }
}