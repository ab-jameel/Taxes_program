Imports System.Data.OleDb
Module Connection_Module
    Public Con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & Application.StartupPath & "\Database.mdb;user id=admin;jet oledb:database password=8383")
End Module
