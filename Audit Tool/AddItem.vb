Imports System.Drawing
Imports System.Windows.Forms
Imports System.Xml
Imports System.Environment

Public Class AddItem
    Private Sub AddItem_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GUISetup()
    End Sub

    Sub GUISetup()
        Me.Text = "Add Branch Code"
        Me.BackColor = Color.White
        Me.FormBorderStyle = FormBorderStyle.FixedToolWindow
        Dim Padding As Integer = 15

        Dim lCode As New Label
        lCode.Location = New Point(Padding, Padding)
        lCode.Text = "Branch Code"
        lCode.BackColor = Color.Transparent
        Me.Controls.Add(lCode)

        Dim Code As New TextBox
        Code.Name = "Code"
        Code.Location = New Point(Padding, lCode.Location.Y + lCode.Height)
        Code.Width = 200
        Code.MaxLength = 3
        Me.Controls.Add(Code)

        Dim lName As New Label
        lName.Location = New Point(Padding, Code.Location.Y + Code.Height + Padding)
        lName.Text = "Branch Name"
        lName.BackColor = Color.Transparent
        Me.Controls.Add(lName)

        Dim Name As New TextBox
        Name.Name = "Name"
        Name.Location = New Point(Padding, lName.Location.Y + lName.Height)
        Name.Width = 200
        Me.Controls.Add(Name)

        Dim Submit As New Button
        Submit.Text = "Add Branch Code"
        Submit.Location = New Point(Padding, Name.Location.Y + Name.Height + Padding)
        Submit.Width = 200
        Me.Controls.Add(Submit)
        AddHandler Submit.Click, AddressOf AddEntry

        Me.Size = New Size(Submit.Width + Padding * 2 + Me.Width - Me.ClientSize.Width, Submit.Location.Y + Submit.Height + Padding + Me.Height - Me.ClientSize.Height)
        Me.Top = (Screen.PrimaryScreen.Bounds.Height - Me.Height) / 2
        Me.Left = (Screen.PrimaryScreen.Bounds.Width - Me.Width) / 2
    End Sub

    Sub AddEntry()
        Dim Bool As Boolean = True
        Dim Code As TextBox = Me.Controls("Code")
        Dim Name As TextBox = Me.Controls("Name")
        For Each Control As Control In Me.Controls
            If TypeOf Control Is TextBox Then
                If Control.Text = String.Empty Then Bool = False
            End If
        Next
        If Bool Then
            Dim XML As String = GetFolderPath(SpecialFolder.ApplicationData) & "\AuditAnalyzer\BranchCodes.XML"
            Dim XMLDoc As New XmlDocument
            XMLDoc.Load(XML)
            Dim root As XmlNode = XMLDoc.DocumentElement
            Dim Element As XmlElement = XMLDoc.CreateElement(Name.Text.Replace(" ", "_").Replace("/", "--").Replace("'", "..").Replace("(", "-.").Replace(")", ".-"))
            Element.InnerText = Code.Text
            root.AppendChild(Element)
            XMLDoc.Save(XML)
            Me.Hide()
            Me.Dispose()
        Else
            MsgBox("Please ensure all fields have been filled.")
        End If
    End Sub
End Class