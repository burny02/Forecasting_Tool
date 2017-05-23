Module SaveModule

    Public Sub Saver(ctl As Object)
        Dim DisplayMessage As Boolean = True

        'Get a generic command list first - Ignore errors (Multi table)
        Dim cb As New OleDb.OleDbCommandBuilder(OverClass.CurrentDataAdapter)

        Try
            OverClass.CurrentDataAdapter.UpdateCommand = cb.GetUpdateCommand()
        Catch
        End Try
        Try
            OverClass.CurrentDataAdapter.InsertCommand = cb.GetInsertCommand()
        Catch
        End Try
        Try
            OverClass.CurrentDataAdapter.DeleteCommand = cb.GetDeleteCommand()
        Catch
        End Try


        'Create and overwrite a custom one if needed (More than 1 table) ...OLEDB Parameters must be added in the order they are used

        Select Case ctl.name

            Case "DataGridView1", "DataGridView2"

                OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE Staff " &
                                                               "SET FName=@P1, SName=@P2, Archived=@P3 WHERE Staff_ID=@P4")

                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.VarChar, 255, "FName")
                    .Add("@P2", OleDb.OleDbType.VarChar, 255, "SName")
                    .Add("@P3", OleDb.OleDbType.Boolean, 255, "Archived")
                    .Add("@P3", OleDb.OleDbType.Double, 255, "Staff_ID")
                End With

                OverClass.CurrentDataAdapter.InsertCommand = New OleDb.OleDbCommand("INSERT INTO Staff (FName, SName) " &
                                                               "VALUES (@P1,@P2)")

                With OverClass.CurrentDataAdapter.InsertCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.VarChar, 255, "FName")
                    .Add("@P2", OleDb.OleDbType.VarChar, 255, "SName")
                End With


        End Select

        Call OverClass.SetCommandConnection()
        Call OverClass.UpdateBackend(ctl, DisplayMessage)
        If DisplayMessage = False Then MsgBox("Table Updated")


    End Sub

End Module
