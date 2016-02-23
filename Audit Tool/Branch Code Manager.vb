Imports System.Drawing
Imports System.Windows.Forms
Imports System.Xml
Imports System.Environment
Imports System.Security.Cryptography

Public Class Branch_Code_Manager
    Private Sub Branch_Code_Manager_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GUISetup()
        AddItems()
    End Sub

    Private Sub Branch_Code_Manager_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        AddItems()
    End Sub

    Sub GUISetup()
        Me.Text = "Branch Code Manager"
        Me.Size = New Size(500, 500)
        Me.Top = (Screen.PrimaryScreen.Bounds.Height - Me.Height) / 2
        Me.Left = (Screen.PrimaryScreen.Bounds.Width - Me.Width) / 2
        Me.BackColor = Color.White
        Me.FormBorderStyle = FormBorderStyle.FixedToolWindow
        Dim Padding As Integer = 25

        Dim Title As New Label
        Title.Text = "Branch Codes"
        Title.Location = New Point(Padding, Padding)
        Me.Controls.Add(Title)

        Dim List As New ListView
        List.Name = "List"
        List.Location = New Point(Title.Location.X, Title.Location.Y + Title.Height)
        List.Size = New Size(Me.ClientSize.Width - List.Location.X - Padding * 2 - 100, 300)
        List.View = View.Details
        Dim ColumnOne As Integer = List.Width * 0.15
        Dim ColumnTwo As Integer = List.Width * 0.85
        List.Columns.Add("Code", ColumnOne, HorizontalAlignment.Left)
        List.Columns.Add("Branch", ColumnTwo, HorizontalAlignment.Left)
        Me.Controls.Add(List)

        Dim Add As New Button
        Add.Text = "Add Entry"
        Add.Size = New Size(100, 25)
        Add.Location = New Point(List.Location.X + List.Width + Padding, List.Location.Y)
        Add.FlatStyle = FlatStyle.Flat
        Add.Enabled = False
        Me.Controls.Add(Add)
        AddHandler Add.Click, AddressOf AddEntry

        Dim Delete As New Button
        Delete.Text = "Delete Entry"
        Delete.Size = New Size(100, 25)
        Delete.Location = New Point(List.Location.X + List.Width + Padding, Add.Location.Y + Add.Height + Padding)
        Delete.FlatStyle = FlatStyle.Flat
        Delete.Enabled = False
        Me.Controls.Add(Delete)
        AddHandler Delete.Click, AddressOf DeleteEntry

        Dim Import As New Button
        Import.Text = "Import"
        Import.Size = New Size(100, 25)
        Import.Location = New Point(List.Location.X + List.Width + Padding, Delete.Location.Y + Delete.Height + Padding)
        Import.FlatStyle = FlatStyle.Flat
        Import.Enabled = False
        Me.Controls.Add(Import)
        AddHandler Import.Click, AddressOf ImportXML

        Dim Export As New Button
        Export.Text = "Export"
        Export.Size = New Size(100, 25)
        Export.Location = New Point(List.Location.X + List.Width + Padding, Import.Location.Y + Import.Height + Padding)
        Export.FlatStyle = FlatStyle.Flat
        Export.Enabled = False
        Me.Controls.Add(Export)
        AddHandler Export.Click, AddressOf ExportXML

        Dim Password As New TextBox
        Password.Name = "Password"
        Password.Location = New Point(List.Location.X, List.Location.Y + List.Height + Padding)
        Password.Multiline = True
        Password.Size = New Size(List.Width, 25)
        Password.PasswordChar = "*"
        Me.Controls.Add(Password)

        Dim ChangePassword As New Label
        ChangePassword.Width = Password.Width
        ChangePassword.Location = New Point(Password.Location.X, Password.Location.Y + Password.Height)
        ChangePassword.Text = "Change Password"
        ChangePassword.Cursor = Cursors.Hand
        ChangePassword.ForeColor = Color.Black
        Me.Controls.Add(ChangePassword)
        AddHandler ChangePassword.Click, AddressOf PasswordChange

        Dim Unlock As New Button
        Unlock.Text = "Unlock"
        Unlock.Size = New Size(100, 25)
        Unlock.Location = New Point(Password.Location.X + Password.Width + Padding, Password.Location.Y)
        Unlock.FlatStyle = FlatStyle.Flat
        Unlock.Cursor = Cursors.Hand
        Me.Controls.Add(Unlock)
        AddHandler Unlock.Click, AddressOf Enable

        Me.Height = Password.Location.Y + Password.Height + Padding + (Me.Height - Me.ClientSize.Height) + ChangePassword.Height
    End Sub

    Sub AddItems()
        Dim XML As String = GetFolderPath(SpecialFolder.ApplicationData) & "\AuditAnalyzer\BranchCodes.XML"
        Dim XMLDoc As New XmlDocument
        Dim BranchCodes(1) As String
        Dim List As ListView = Me.Controls("List")
        List.Items.Clear()
        XMLDoc.Load(XML)
        Dim Node As XmlNode
        For Each Node In XMLDoc.DocumentElement.ChildNodes
            BranchCodes(1) = Node.Name.Replace("_", " ").Replace("--", "/").Replace("..", "'").Replace(".-", ")").Replace("-.", "(")
            BranchCodes(0) = Node.InnerText
            List.Items.Add(New ListViewItem(BranchCodes))
        Next
        List.Sort()
    End Sub

    Sub AddEntry()
        Dim Add As Form = New AddItem
        Add.ShowDialog()
    End Sub

    Sub DeleteEntry()
        Dim XML As String = GetFolderPath(SpecialFolder.ApplicationData) & "\AuditAnalyzer\BranchCodes.XML"
        Dim XMLDoc As New XmlDocument
        Dim List As ListView = Me.Controls("List")
        XMLDoc.Load(XML)
        For Each Item As ListViewItem In List.SelectedItems
            If MsgBox("Are you sure you want to remove Branch Code " & Item.Text & "?", MsgBoxStyle.OkCancel) = DialogResult.OK Then
                For Each Node As XmlNode In XMLDoc.DocumentElement.ChildNodes
                    If Node.Name = Item.SubItems(1).Text Then
                        Node.ParentNode.RemoveChild(Node)
                        XMLDoc.Save(XML)
                        AddItems()
                    End If
                Next
            End If
        Next
    End Sub

    Sub ImportXML()
        Dim Import As New OpenFileDialog
        Import.Filter = "XML Document|*.xml"
        Import.Title = "Import Branch Codes"
        If Import.ShowDialog = DialogResult.OK Then
            Dim Bool As Boolean = True
            Dim XML As String = GetFolderPath(SpecialFolder.ApplicationData) & "\AuditAnalyzer\BranchCodes.XML"
            Dim XMLDoc As New XmlDocument
            Dim cXMLDoc As New XmlDocument
            Dim List As ListView = Me.Controls("List")
            XMLDoc.Load(XML)
            cXMLDoc.Load(Import.FileName)
            Dim root As XmlNode = XMLDoc.DocumentElement
            Dim Element As XmlElement
            For Each cNode As XmlNode In cXMLDoc.DocumentElement.ChildNodes
                Bool = True
                For Each Node As XmlNode In XMLDoc.DocumentElement.ChildNodes
                    If Node.Name = Node.Name Then Bool = False
                Next
                Element = XMLDoc.CreateElement(cNode.Name)
                Element.InnerText = cNode.InnerText
                root.AppendChild(Element)
                XMLDoc.Save(XML)
            Next
            AddItems()
        End If
    End Sub

    Sub ExportXML()
        Dim XML As String = GetFolderPath(SpecialFolder.ApplicationData) & "\AuditAnalyzer\BranchCodes.XML"
        Dim Save As New SaveFileDialog
        Save.Filter = "XML Document|*.xml"
        Save.Title = "Export Branch Codes"
        If Save.ShowDialog = DialogResult.OK Then
            My.Computer.FileSystem.CopyFile(XML, Save.FileName, True)
        End If
    End Sub

    Sub Enable()
        Dim PasswordText As TextBox = Me.Controls("Password")
        If Hash(PasswordText.Text) = Password() Then
            For Each Control As Control In Me.Controls
                If TypeOf Control Is Button Then
                    Control.Enabled = True
                    Control.Cursor = Cursors.Hand
                End If
            Next
        Else
            MsgBox("Incorrect password. Please try again.")
        End If
    End Sub

    Function Hash(ByVal sSourceData As String) As String
        Dim tmpSource() As Byte = ASCIIEncoding.ASCII.GetBytes(sSourceData)
        Dim HashByte() As Byte = New MD5CryptoServiceProvider().ComputeHash(tmpSource)
        Hash = ByteArrayToString(HashByte)
    End Function

    Private Function ByteArrayToString(ByVal arrInput() As Byte) As String
        Dim i As Integer
        Dim sOutput As New StringBuilder(arrInput.Length)
        For i = 0 To arrInput.Length - 1
            sOutput.Append(arrInput(i).ToString("X2"))
        Next
        Return sOutput.ToString()
    End Function

    Sub PasswordChange()
        Dim Password As Form = New Password
        Password.ShowDialog()
    End Sub

    Function Password(Optional ByVal NewPassword As String = "") As String
        Password = String.Empty
        Dim XMLDir As String = GetFolderPath(SpecialFolder.ApplicationData) & "\AuditAnalyzer"
        Dim XML As String = XMLDir & "\Password.XML"
        If Not My.Computer.FileSystem.DirectoryExists(XMLDir) Then
            My.Computer.FileSystem.CreateDirectory(XMLDir)
        End If
        If Not My.Computer.FileSystem.FileExists(XML) Then
            Dim settings As XmlWriterSettings = New XmlWriterSettings()
            settings.Indent = True

            Using writer As XmlWriter = XmlWriter.Create(XML, settings)
                writer.WriteStartDocument()
                writer.WriteStartElement("Settings")
                writer.WriteStartElement("Password")
                writer.WriteEndElement()
                writer.WriteEndElement()
                writer.WriteEndDocument()
            End Using
        End If
        Dim XMLDoc As New XmlDocument
        XMLDoc.Load(XML)
        For Each Node As XmlNode In XMLDoc.DocumentElement.ChildNodes
            If Not String.IsNullOrEmpty(NewPassword) Then Node.InnerText = NewPassword
            Password = Node.InnerText
            XMLDoc.Save(XML)
        Next
    End Function
End Class