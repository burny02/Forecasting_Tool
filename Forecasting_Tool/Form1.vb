Imports Microsoft.Reporting.WinForms

Public Class Form1
    Private SpreadsheetDate As Date
    Private CurrentTable As String
    Private TableDivision As String
    Public Source As String = "Link"
    Public Scenario As Boolean = False

    Public Sub New(WhatSource As String, WhatScenario As Boolean)
        Source = WhatSource
        Scenario = WhatScenario
        InitializeComponent()
        FormCol.Add(Me)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Scenario = False Then

            Me.WindowState = FormWindowState.Maximized

            Call StartUp(Me)

            Try
                Me.Label2.Text = SolutionName & vbNewLine & "Developed by David Burnside" & vbNewLine & "Version: " & System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
            Catch
                Me.Label2.Text = SolutionName & vbNewLine & "Developed by David Burnside"
            End Try

            Me.Text = SolutionName

        End If

    End Sub

    Public Sub TabControl1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl1.Selecting


        Select Case e.TabPage.Text

            Case "Assumptions"
                If ComboBox1.Items.Count = 0 Then
                    ComboBox1.Items.Add("hSite Screening")
                    ComboBox1.Items.Add("hSite Quarantine")
                    ComboBox1.Items.Add("hSite Permanent Headcount")
                    ComboBox1.Items.Add("Finance")
                End If
                If ComboBox4.Items.Count = 0 Then
                    ComboBox4.Items.Add("hSite Bank Cost")
                    ComboBox4.Items.Add("hSite Bank Hours")
                    ComboBox4.Items.Add("hSite Bank Rota")
                End If
                If ComboBox2.Items.Count = 0 Then
                    ComboBox2.Items.Add("Raw Data")
                    ComboBox2.Items.Add("Graph")
                End If
                ComboBox2.SelectedIndex = ComboBox2.FindStringExact("Graph")
                ComboBox4.SelectedIndex = ComboBox4.FindStringExact("hSite Bank Hours")
                ComboBox1.SelectedIndex = ComboBox1.FindStringExact("hSite Screening")
        End Select
    End Sub


    Private Sub ComboBox1_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        If Scenario = True Then Exit Sub
        DataGridView1.Columns.Clear()
        Dim SqlCode As String = vbNullString

        Select Case ComboBox1.Items(ComboBox1.SelectedIndex)

            Case "hSite Screening"
                SqlCode = "SELECT * FROM hSite_Assumptions ORDER BY ID DESC"
                CurrentTable = "hSite_Assumptions"
            Case "hSite Quarantine"
                SqlCode = "SELECT * FROM Q_Group ORDER BY ID DESC"
                CurrentTable = "Q_Group"
            Case "Finance"
                SqlCode = "SELECT * FROM Finance_Assumptions ORDER BY ID Desc"
                CurrentTable = "Finance_Assumptions"
            Case "hSite Permanent Headcount"
                SqlCode = "SELECT * FROM hSite_Perm ORDER BY ID Desc"
                CurrentTable = "hSite_Perm"

        End Select


        If SqlCode <> vbNullString Then
            OverClass.CreateDataSet(SqlCode, BindingSource1, DataGridView1)
            DataGridView1.Columns("ID").Visible = False
        End If

        SetDivision()
        If Scenario = False Then UpdateFromSpreadsheet()

        Select Case ComboBox1.Items(ComboBox1.SelectedIndex)

            Case "hSite Screening"
            Case "hSite Quarantine"
                Dim ViewClm As New DataGridViewImageColumn
                ViewClm.HeaderText = "View Group"
                ViewClm.Name = "ViewClm"
                ViewClm.ImageLayout = DataGridViewImageCellLayout.Zoom
                ViewClm.Image = My.Resources.Magnifying_Glass
                DataGridView1.Columns.Add(ViewClm)

        End Select

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        RefreshReport(False)
    End Sub

    Private Sub RefreshReport(UpdateAll As Boolean)

        For Each form As Form1 In FormCol
            If UpdateAll = False And form IsNot Me Then Continue For
            Dim SQLQuery As String = vbNullString
            If String.IsNullOrEmpty(form.ComboBox4.Text) Then Exit Sub
            If String.IsNullOrEmpty(form.ComboBox2.Text) Then Exit Sub
            form.ReportViewer1.Visible = False
            form.DataGridView2.Visible = False
            form.DataGridView2.Columns.Clear()
            form.ReportViewer1.Clear()
            form.ReportViewer1.LocalReport.DataSources.Clear()
            form.SpreadsheetDate = IO.File.GetLastWriteTime("M:\VOLUNTEER SCREENING SERVICES\Systems\Forecasting_Tool\OpsToolOutput.xlsx")
            form.Label5.Text = "Based on spreadsheet dated: " & Strings.Format(SpreadsheetDate, "dd-MMM-yyyy HH:mm")

            If form.ComboBox2.Items(form.ComboBox2.SelectedIndex) = "Graph" Then
                form.ReportViewer1.Visible = True
                form.ReportViewer1.LocalReport.ReportEmbeddedResource = "Forecasting_Tool." & form.ComboBox4.Items(form.ComboBox4.SelectedIndex) & ".rdlc"
                Select Case form.ComboBox4.Items(form.ComboBox4.SelectedIndex)

                    Case "hSite Bank Cost"
                        SQLQuery = "SELECT DateField, P_Price, N_Price, C_Price FROM " & hSite_Bank_Per_Month(form.Source) & " WHERE cdate(format(Datefield,'MMM yyyy'))>=cdate(format(date(),'MMM yyyy'))"
                    Case "hSite Bank Hours"
                        SQLQuery = "SELECT Datefield, P_Hours, N_Hours, C_Hours FROM " & hSite_Bank_Per_Month(form.Source) & " WHERE cdate(format(Datefield,'MMM yyyy'))>=cdate(format(date(),'MMM yyyy'))"
                    Case "hSite Bank Rota"
                        SQLQuery = "SELECT * FROM " & hSite_Bank_Per_Week(form.Source) & " WHERE cdate(format(MinOfWhatDay,'MMM yyyy')) BETWEEN cdate(format(date(),'MMM yyyy')) AND dateadd('m',1,cdate(format(date(),'MMM yyyy')))"

                End Select

                form.ReportViewer1.LocalReport.DataSources.Add(New ReportDataSource("ReportDataSet",
                                                      OverClass.TempDataTable(SQLQuery)))
                form.ReportViewer1.RefreshReport()
            Else
                form.DataGridView2.Visible = True
                Select Case form.ComboBox4.Items(form.ComboBox4.SelectedIndex)

                    Case "hSite Bank Cost"
                        SQLQuery = "SELECT WhatMonth, P_Price, N_Price, C_Price FROM " & hSite_Bank_Per_Month(form.Source) & " WHERE cdate(format(Datefield,'MMM yyyy'))>=cdate(format(date(),'MMM yyyy')) ORDER BY DateField"
                    Case "hSite Bank Hours"
                        SQLQuery = "SELECT WhatMonth, P_Hours, N_Hours, C_Hours FROM " & hSite_Bank_Per_Month(form.Source) & " WHERE cdate(format(Datefield,'MMM yyyy'))>=cdate(format(date(),'MMM yyyy')) ORDER BY DateField"
                    Case "hSite Bank Rota"
                        SQLQuery = "SELECT * FROM " & hSite_Bank_Per_Week(form.Source) & " WHERE cdate(format(MinOfWhatDay,'MMM yyyy'))>=cdate(format(date(),'MMM yyyy')) ORDER BY MinOfWhatDay"

                End Select
                form.DataGridView2.DataSource = OverClass.TempDataTable(SQLQuery)
            End If
        Next

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        RefreshReport(False)
    End Sub

    Private Sub DataGridView1_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        If e.ColumnIndex = DataGridView1.Columns("Selected").Index Then
            If (Role <> TableDivision) Then
                MsgBox("You do not have permission to change this divisions assumptions.")
                Exit Sub
            End If

            If MsgBox("Do you want to select and use this assumption for all calculations?", vbYesNo) = vbYes Then
                OverClass.AddToMassSQL("UPDATE " & CurrentTable & " SET SELECTED=False")
                OverClass.AddToMassSQL("UPDATE " & CurrentTable & " SET SELECTED=TRUE WHERE ID=" & DataGridView1.Item("ID", e.RowIndex).Value)
                OverClass.ExecuteMassSQL()
                ComboBox1_SelectedIndexChanged_1(Me, New EventArgs)
                RefreshReport(True)
                Exit Sub
            End If
        End If
        If e.ColumnIndex = DataGridView1.Columns("ViewClm").Index Then
            Dim GroupID As Long = DataGridView1.Item("ID", e.RowIndex).Value
            DataGridView1.Columns.Clear()
            OverClass.CreateDataSet("SELECT * FROM hSite_Q_Assumptions WHERE GroupID=" & GroupID, BindingSource1, DataGridView1)
            ComboBox1.Text = "HSite_Q_Assumptions; Group" & GroupID
            DataGridView1.Columns("ID").Visible = False
            DataGridView1.Columns("GroupID").Visible = False
            Exit Sub
        End If
    End Sub


    Private Sub SetDivision()

        Select Case CurrentTable

            Case "hSite_Assumptions"
                TableDivision = "hSite"
            Case "Q_Group"
                TableDivision = "hSite"
            Case "hSite Permanent Headcount"
                TableDivision = "hSite"
            Case "Finance_Assumptions"
                TableDivision = "Finance"


        End Select

    End Sub

    Private Sub UpdateFromSpreadsheet()
        Dim Update As Boolean = False
        Dim dt As DataTable = OverClass.TempDataTable("SELECT UpdateDate FROM SpreadsheetDate")
        If dt.Rows.Count = 0 Then
            Update = True
        Else
            If (Strings.Format(dt.Rows(0).Item(0), "dd-MMM-yyyy HH:mm") <> Strings.Format(SpreadsheetDate, "dd-MMM-yyyy HH:mm")) Then Update = True
        End If

        If Update = True Then
            OverClass.AddToMassSQL("UPDATE DaysLink INNER JOIN Link ON DaysLink.WhatDay=Link.WhatDay SET DaysLink.PSP=Link.PSP, DaysLink.SSS=Link.SSS, " &
                              "DaysLink.PSP2=Link.PSP2, DaysLink.FUP=Link.FUP, DaysLink.QRisk=Link.QRisk", False)
            OverClass.AddToMassSQL("DELETE * FROM SpreadsheetDate", False)
            OverClass.AddToMassSQL("INSERT INTO SpreadsheetDate (UpdateDate) VALUES (#" & Strings.Format(SpreadsheetDate, "dd-MMM-yyyy HH:mm") & "#)", False)
            OverClass.AddToMassSQL("INSERT INTO History (WhatDay,PSP,SSS,PSP2,FUP,QRisk,SpreadsheetDate) SELECT WhatDay, PSP, SSS, PSP2, FUP, QRisk," &
                                   "#" & Strings.Format(SpreadsheetDate, "dd-MMM-yyyy HH:mm") & "# FROM Link WHERE isnull(WhatDay)=false")
            OverClass.ExecuteMassSQL()
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Variables.LinkNewTable()
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        If Scenario = True Then
            OverClass.AddToMassSQL("DROP TABLE " & Source)
        Else
            OverClass.ExecuteMassSQL()
        End If

    End Sub
End Class
