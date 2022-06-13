Imports System.Data.Odbc
Module Module1
    Public conn As OdbcConnection
    Public cmd As OdbcCommand
    Public rd As OdbcDataReader
    Public da As OdbcDataAdapter
    Public ds As DataSet

    Sub connection()
        conn = New OdbcConnection("dsn=dsn_windowsapplication1")
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
    End Sub
End Module
