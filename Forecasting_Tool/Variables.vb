Imports TemplateDB

Module Variables
    Public OverClass As OverClass
    Private Const TablePath As String = "M:\VOLUNTEER SCREENING SERVICES\Systems\Forecasting_Tool\Forecast Tool.accdb"
    Private Const PWord As String = "RetroRetro*1"
    Private Const Connect2 As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TablePath & ";Jet OLEDB:Database Password=" & PWord
    Private Const UserTable As String = "[Users]"
    Private Const UserField As String = "Username"
    Private Const AuditTbl As String = "[Audit]"
    Private Contact As String = "Ivor Pegington"
    Public Const SolutionName As String = "Forecasting Tool"
    Public Role As String = vbNullString
    Public FormCol As New Collection



    Public Function GetTheConnection() As String
        GetTheConnection = Connect2
    End Function


    Public Sub StartUp(WhatForm As Form)

        OverClass = New TemplateDB.OverClass
        OverClass.SetPrivate(UserTable,
                           UserField,
                           Contact,
                           Connect2,
                           AuditTbl)

        Dim SQLString(0) As String
        SQLString(0) = "SELECT Role FROM [Users] WHERE UserName='" & OverClass.GetUserName & "'"

        Dim dt() As DataTable = OverClass.LoginCheck(SQLString)

        Role = dt(1).Rows(0).Item(0)

        OverClass.LoginCheck()

        OverClass.AddAllDataItem(WhatForm)

        For Each ctl In OverClass.DataItemCollection
            If (TypeOf ctl Is Button) Then
                Dim But As Button = ctl
                AddHandler But.Click, AddressOf ButtonSpecifics
            End If
        Next


    End Sub


    Public Sub LinkNewTable()

        Dim filnam As String = vbNullString
        Dim concon = New ADODB.Connection

        Try
            Dim fd As OpenFileDialog = New OpenFileDialog()
            Dim ColumnHeaderNumber As Integer = 2

            fd.Title = "Choose Excel source"
            fd.Filter = "Excel Files|*.xls;*.xlsx"
            fd.FilterIndex = 1
            fd.RestoreDirectory = True
            fd.Multiselect = False
            fd.AutoUpgradeEnabled = False

            If fd.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
                fd = Nothing
                Exit Sub
            End If

            filnam = fd.FileName
            fd = Nothing

            If Not filnam = vbNullString Then
                Try
                    Dim ADOXTable As New ADOX.Table
                    Dim ADOXCatalog As New ADOX.Catalog
                    Dim TableName = Replace(OverClass.GetUserName, ".", "") & Format(Now(), "ddmmyyyyHHmmss")
                    concon.ConnectionString = Connect2
                    concon.Open()

                    ADOXCatalog.ActiveConnection = concon

                    ADOXTable.ParentCatalog = ADOXCatalog

                    With ADOXTable
                        .Name = TableName
                        .Properties("Jet OLEDB:Link Provider String").Value = "Excel 8.0;DATABASE=" & filnam & ";HDR=Yes"
                        .Properties("Jet OLEDB:Remote Table Name").Value = "Output for Resourcing$"
                        .Properties("Jet OLEDB:Create Link").Value = True
                    End With

                    ADOXCatalog.Tables.Append(ADOXTable)
                    Dim NewFrm As New Form1(TableName, True)
                    NewFrm.Text = "SCENARIO: " & filnam
                    NewFrm.TabControl1.SelectedIndex = 1
                    NewFrm.TabControl1_Selecting(NewFrm.TabControl1, New TabControlCancelEventArgs(NewFrm.TabPage2, 0, False, TabControlAction.Selecting))
                    NewFrm.TabControl1.TabPages.Remove(NewFrm.TabControl1.TabPages(0))
                    NewFrm.Button1.Visible = False
                    NewFrm.Label5.Visible = False
                    NewFrm.ComboBox1.Visible = False
                    NewFrm.Label1.Visible = False
                    NewFrm.DataGridView1.Visible = False
                    NewFrm.TableLayoutPanel1.RowStyles.Item(0).Height = 0
                    NewFrm.TableLayoutPanel1.RowStyles.Item(1).Height = 0
                    NewFrm.Show()


                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                Finally
                    concon.Close()
                    concon = Nothing
                End Try
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

End Module
