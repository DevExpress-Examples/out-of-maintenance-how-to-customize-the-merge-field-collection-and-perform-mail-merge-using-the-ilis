Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Windows
Imports DevExpress.Xpf.Core
Imports DevExpress.Xpf.Ribbon
Imports System.ComponentModel
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.Mvvm

Namespace SimpleDataMerge

    Public Partial Class MainWindow
        Inherits DXRibbonWindow

        Private mergeFieldsNamesInfo As MergeFieldNameInfo()

        Public Sub New()
            Me.InitializeComponent()
            mergeFieldsNamesInfo = TypeDescriptor.GetProperties(GetType(Employee)).Cast(Of PropertyDescriptor)().[Select](Function([property]) New MergeFieldNameInfo([property].Name, [property].DisplayName)).ToArray()
        End Sub

        Private Sub OnRichEditCustomizeMergeFields(ByVal sender As Object, ByVal e As DevExpress.XtraRichEdit.CustomizeMergeFieldsEventArgs)
            e.MergeFieldsNames = mergeFieldsNamesInfo.Where(Function(info) info.CanShow).[Select](Function(info) New MergeFieldName(info.Name, String.Format("{0} ({1})", info.DisplayName, info.Name))).ToArray()
        End Sub

        Private Sub OnCustomizeMergeFields(ByVal sender As Object, ByVal e As DevExpress.Xpf.Bars.ItemClickEventArgs)
            Dim control As CustomizeMergeFieldsControl = New CustomizeMergeFieldsControl(mergeFieldsNamesInfo)
            Dim window As DXDialogWindow = New DXDialogWindow()
            window.Content = control
            window.Title = "Customize merge fields"
            window.Width = 400
            window.Height = 400
            window.FooterButtons.Add(New DialogButton() With {.Content = "OK", .IsDefault = True, .DialogResult = MessageResult.OK})
            window.Owner = GetWindow(Me)
            window.WindowStartupLocation = WindowStartupLocation.CenterOwner
            window.WindowStyle = WindowStyle.ToolWindow
            window.ShowDialog()
        End Sub
    End Class

    Public Module EmployeeData

        Private firstName As String() = {"Nancy", "Andrew", "Janet", "Margaret", "Steven", "Michael", "Robert", "Laura", "Anne"}

        Private lastName As String() = {"Davolio", "Fuller", "Leverling", "Peacock", "Buchanan", "Suyama", "King", "Callahan", "Dodsworth"}

        Private city As String() = {"Seattle", "Tacoma", "Kirkland", "Redmond", "London", "London", "London", "Seattle", "London"}

        Private country As String() = {"USA", "USA", "USA", "USA", "UK", "UK", "UK", "USA", "UK"}

        Private address As String() = {"507 - 20th Ave. E. Apt. 2A", "908 W. Capital Way", "722 Moss Bay Blvd.", "4110 Old Redmond Rd.", "14 Garrett Hill", "Coventry House Miner Rd.", "Edgeham Hollow Winchester Way", "4726 - 11th Ave. N.E.", "7 Houndstooth Rd."}

        Private position As String() = {"Sales Representative", "Vice President, Sales", "Sales Representative", "Sales Representative", "Sales Manager", "Sales Representative", "Sales Representative", "Inside Sales Coordinator", "Sales Representative"}

        Private gender As Char() = {"F"c, "M"c, "F"c, "F"c, "M"c, "M"c, "M"c, "F"c, "F"c}

        Private phone As String() = {"(206) 555-9857", "(206) 555-9482", "(206) 555-3412", "(206) 555-8122", "(71) 555-4848", "(71) 555-7773", "(71) 555-5598", "(206) 555-1189", "(71) 555-4444"}

        Private companyName As String() = {"Consolidated Holdings", "Around the Horn", "North/South", "Island Trading", "White Clover Markets", "Trail's Head Gourmet Provisioners", "The Cracker Box", "The Big Cheese", "Rattlesnake Canyon Grocery", "Split Rail Beer & Ale", "Hungry Coyote Import Store", "Great Lakes Food Market"}

        Public ReadOnly Property Employees As List(Of Employee)
            Get
                Return Enumerable.Range(0, 10).[Select](New Func(Of Integer, Employee)(AddressOf CreateEmployee)).ToList()
            End Get
        End Property

        Private Function CreateEmployee(ByVal seed As Integer) As Employee
            Dim result As Employee = New Employee()
            Dim rnd As Random = New Random(seed)
            Dim countryIndex As Integer = rnd.Next(0, country.Length)
            result.Country = country(countryIndex)
            result.Address = address(countryIndex)
            result.City = city(countryIndex)
            result.LastName = lastName(rnd.Next(0, lastName.Length))
            Dim firstNameIndex As Integer = rnd.Next(0, firstName.Length)
            result.FirstName = firstName(firstNameIndex)
            result.Gender = gender(firstNameIndex)
            result.HiringDate = Date.Now.AddDays(-rnd.Next(0, 2000))
            result.Position = position(rnd.Next(0, position.Length))
            result.Phone = phone(rnd.Next(0, phone.Length))
            result.CompanyName = companyName(rnd.Next(0, companyName.Length))
            result.HRManagerName = String.Format("{0} {1}", firstName(rnd.Next(0, firstName.Length)), lastName(rnd.Next(0, lastName.Length)))
            Return result
        End Function
    End Module

    Public Class Employee

        <DisplayName("First Name")>
        Public Property FirstName As String

        <DisplayName("Last Name")>
        Public Property LastName As String

        <DisplayName("Hiring Date")>
        Public Property HiringDate As Date

        Public Property Address As String

        Public Property City As String

        Public Property Country As String

        Public Property Position As String

        <DisplayName("Company Name")>
        Public Property CompanyName As String

        Public Property Gender As Char

        Public Property Phone As String

        <DisplayName("HR Manager Name")>
        Public Property HRManagerName As String
    End Class

    Public Class MergeFieldNameInfo

        Public Sub New(ByVal name As String, ByVal displayName As String)
            Me.Name = name
            Me.DisplayName = displayName
            CanShow = True
        End Sub

        Public Property Name As String

        Public Property DisplayName As String

        Public Property CanShow As Boolean
    End Class
End Namespace
