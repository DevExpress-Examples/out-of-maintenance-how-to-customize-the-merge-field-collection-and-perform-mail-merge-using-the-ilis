using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using DevExpress.Xpf.Core;
using DevExpress.Xpf.Ribbon;
using System.ComponentModel;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.Mvvm;

namespace SimpleDataMerge
{
    public partial class MainWindow : DXRibbonWindow
    {
        MergeFieldNameInfo[] mergeFieldsNamesInfo;
        public MainWindow()
        {
            InitializeComponent();
            this.mergeFieldsNamesInfo = TypeDescriptor.GetProperties(typeof(Employee))
    .Cast<PropertyDescriptor>()
    .Select(property => new MergeFieldNameInfo(property.Name, property.DisplayName))
    .ToArray();
        }


        private void OnRichEditCustomizeMergeFields(object sender, DevExpress.XtraRichEdit.CustomizeMergeFieldsEventArgs e)
        {
            e.MergeFieldsNames = mergeFieldsNamesInfo
                .Where(info => info.CanShow)
                .Select(info => new MergeFieldName(info.Name, string.Format("{0} ({1})", info.DisplayName, info.Name)))
                .ToArray();
        }
        void OnCustomizeMergeFields(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            CustomizeMergeFieldsControl control = new CustomizeMergeFieldsControl(mergeFieldsNamesInfo);

            DXDialogWindow window = new DXDialogWindow();
            window.Content = control;
            window.Title = "Customize merge fields";
            window.Width = 400;
            window.Height = 400;
            window.FooterButtons.Add(new DialogButton() { Content = "OK", IsDefault = true, DialogResult = MessageResult.OK });
            window.Owner = Window.GetWindow(this);
            window.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            window.WindowStyle = WindowStyle.ToolWindow;

            window.ShowDialog();
        }
    }
    public static class EmployeeData
    {
        static string[] firstName = { "Nancy", "Andrew", "Janet", "Margaret", "Steven", "Michael", "Robert", "Laura", "Anne" };
        static string[] lastName = { "Davolio", "Fuller", "Leverling", "Peacock", "Buchanan", "Suyama", "King", "Callahan", "Dodsworth" };
        static string[] city = { "Seattle", "Tacoma", "Kirkland", "Redmond", "London", "London", "London", "Seattle", "London" };
        static string[] country = { "USA", "USA", "USA", "USA", "UK", "UK", "UK", "USA", "UK" };
        static string[] address = { "507 - 20th Ave. E. Apt. 2A", "908 W. Capital Way", "722 Moss Bay Blvd.", "4110 Old Redmond Rd.", "14 Garrett Hill", "Coventry House Miner Rd.", "Edgeham Hollow Winchester Way", "4726 - 11th Ave. N.E.", "7 Houndstooth Rd." };
        static string[] position = { "Sales Representative", "Vice President, Sales", "Sales Representative", "Sales Representative", "Sales Manager", "Sales Representative", "Sales Representative", "Inside Sales Coordinator", "Sales Representative" };
        static char[] gender = { 'F', 'M', 'F', 'F', 'M', 'M', 'M', 'F', 'F' };
        static string[] phone = { "(206) 555-9857", "(206) 555-9482", "(206) 555-3412", "(206) 555-8122", "(71) 555-4848", "(71) 555-7773", "(71) 555-5598", "(206) 555-1189", "(71) 555-4444" };
        static string[] companyName = { "Consolidated Holdings", "Around the Horn", "North/South", "Island Trading", "White Clover Markets", "Trail's Head Gourmet Provisioners", "The Cracker Box", "The Big Cheese", "Rattlesnake Canyon Grocery", "Split Rail Beer & Ale", "Hungry Coyote Import Store", "Great Lakes Food Market" };
        public static List<Employee> Employees
        {
            get { return Enumerable.Range(0, 10).Select(CreateEmployee).ToList(); }
        }
        static Employee CreateEmployee(int seed)
        {
            Employee result = new Employee();
            Random rnd = new Random(seed);
            int countryIndex = rnd.Next(0, country.Length);
            result.Country = country[countryIndex];
            result.Address = address[countryIndex];
            result.City = city[countryIndex];
            result.LastName = lastName[rnd.Next(0, lastName.Length)];
            int firstNameIndex = rnd.Next(0, firstName.Length);
            result.FirstName = firstName[firstNameIndex];
            result.Gender = gender[firstNameIndex];
            result.HiringDate = DateTime.Now.AddDays(-(rnd.Next(0, 2000)));
            result.Position = position[rnd.Next(0, position.Length)];
            result.Phone = phone[rnd.Next(0, phone.Length)];
            result.CompanyName = companyName[rnd.Next(0, companyName.Length)];
            result.HRManagerName = String.Format("{0} {1}", firstName[rnd.Next(0, firstName.Length)], lastName[rnd.Next(0, lastName.Length)]);
            return result;
        }
    }
    public class Employee
    {
        [DisplayName("First Name")]
        public string FirstName { get; set; }
        [DisplayName("Last Name")]
        public string LastName { get; set; }
        [DisplayName("Hiring Date")]
        public DateTime HiringDate { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string Country { get; set; }
        public string Position { get; set; }
        [DisplayName("Company Name")]
        public string CompanyName { get; set; }
        public char Gender { get; set; }
        public string Phone { get; set; }
        [DisplayName("HR Manager Name")]
        public string HRManagerName { get; set; }
    }
    public class MergeFieldNameInfo
    {
        public MergeFieldNameInfo(string name, string displayName)
        {
            Name = name;
            DisplayName = displayName;
            CanShow = true;
        }
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public bool CanShow { get; set; }
    }


}
