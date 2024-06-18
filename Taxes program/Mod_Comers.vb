Imports System.Data.OleDb
Module Mod_Comers

    Public Function Count_Come(ByVal ID As Integer)
        Dim x As Integer
        Try
            Dim v As New OleDbCommand("select count (ID) from Comers where ID=@ID", Con)
            v.Parameters.Add("ID", OleDbType.Integer).Value = ID
            Con.Open()
            x = v.ExecuteScalar
            Con.Close()
        Catch ex As Exception
            x = 0
            Con.Close()
        End Try
        Return x
    End Function

    Public Sub Insert_Come(ByVal Comer_Name1 As String, ByVal Num_VAT As Double)
        Dim n As New OleDb.OleDbCommand("insert into Comers(Comer_Name1, Num_VAT) values (@Comer_Name1, @Num_VAT)", Con)
        n.Parameters.Add("Comer_Name1", OleDbType.VarChar).Value = Comer_Name1
        n.Parameters.Add("Num_VAT", OleDbType.Double).Value = Num_VAT
        Con.Open()
        n.ExecuteNonQuery()
        Con.Close()
        n = Nothing
    End Sub

    Public Sub Update_Come(ByVal Comer_Name1 As String, ByVal Num_VAT As Double)
        Dim v As New OleDbCommand("update Comers set Comer_Name1=@Comer_Name1,Num_VAT=@Num_VAT where ID = " & Fr_CN_Comers.NC, Con)
        v.Parameters.Add("Comer_Name1", OleDbType.VarChar).Value = Comer_Name1
        v.Parameters.Add("Num_VAT", OleDbType.Double).Value = Num_VAT
        Con.Open()
        v.ExecuteNonQuery()
        Con.Close()
        v = Nothing
    End Sub

    Public Sub Delete_Come(ByVal ID2 As Integer)
        Dim b As New OleDbCommand("Delete from Comers where ID=@ID2", Con)
        b.Parameters.Add("ID2", OleDbType.Integer).Value = ID2
        Con.Open()
        b.ExecuteNonQuery()
        Con.Close()
        b = Nothing
    End Sub

End Module
