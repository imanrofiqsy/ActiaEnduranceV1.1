Imports System.Data.SqlClient
Module DatabaseConnection
    Public Connection As SqlConnection
    Dim dbHostname As String = "DESKTOP-6LJRE3H"
    Dim dbDatabase As String = "db_Endurance_Tester"
    Dim dbUsername As String = "smt2024"
    Dim dbPassword As String = "Pass1234@"

    Public Sub Connect()
        Try
            Dim database As String
            database = "Data Source=" & dbHostname & "\SQLEXPRESS;
                initial catalog=" & dbDatabase & ";
                User ID=" & dbUsername & ";
                Password=" & dbPassword & ";
                MultipleActiveResultSets=true"
            Connection = New SqlConnection(database)
            If Connection.State = ConnectionState.Closed Then Connection.Open() ' Else Connection.Close()
        Catch ex As Exception
            MsgBox("Database Connection Error -> " + ex.Message)
        End Try
    End Sub
End Module
