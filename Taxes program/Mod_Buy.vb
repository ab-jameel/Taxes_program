Imports System.Data.OleDb
Module Mod_Buy

    Public Function Count_Buy(ByVal ID As Integer)
        Dim x As Integer
        Try
            Dim v As New OleDbCommand("select count (ID) from The_Table where ID=@ID", Con)
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

    Public Sub Insert_Buy(ByVal Num As Integer, ByVal Date1 As Date, ByVal Describe1 As String, ByVal Price As Double, ByVal VAT As Double, ByVal Type As String, ByVal Comer_Name As String)
        Dim n As New OleDb.OleDbCommand("insert into The_Table(Num, Date1, Describe1, Price, VAT, Type, Comer_Name) values (@Num, @Date1, @Describe1, @Price, @VAT, @Type, @Comer_Name)", Con)
        n.Parameters.Add("Num", OleDbType.Integer).Value = Num
        n.Parameters.Add("Date1", OleDbType.Date).Value = Date1
        n.Parameters.Add("Describe1", OleDbType.VarChar).Value = Describe1
        n.Parameters.Add("Price", OleDbType.Double).Value = Price
        n.Parameters.Add("VAT", OleDbType.Double).Value = VAT
        n.Parameters.Add("Type", OleDbType.VarChar).Value = Type
        n.Parameters.Add("Comer_Name", OleDbType.VarChar).Value = Comer_Name
        Con.Open()
        n.ExecuteNonQuery()
        Con.Close()
        n = Nothing
    End Sub

    Public Sub Update_Buy(ByVal Num As Integer, ByVal Date1 As Date, ByVal Describe1 As String, ByVal Price As Double, ByVal VAT As Double, ByVal Type As String, ByVal Comer_Name As Integer)
        Dim v As New OleDbCommand("update The_Table set Num=@Num,Date1=@Date1,Describe1=@Describe1,Price=@Price,VAT=@VAT,Type=@Type,Comer_Name=@Comer_Name where ID = " & Buy_Form.NBR, Con)
        v.Parameters.Add("Num", OleDbType.Integer).Value = Num
        v.Parameters.Add("Date1", OleDbType.Date).Value = Date1
        v.Parameters.Add("Describe1", OleDbType.VarChar).Value = Describe1
        v.Parameters.Add("Price", OleDbType.Double).Value = Price
        v.Parameters.Add("VAT", OleDbType.Double).Value = VAT
        v.Parameters.Add("Type", OleDbType.VarChar).Value = Type
        v.Parameters.Add("Comer_Name", OleDbType.Integer).Value = Comer_Name
        Con.Open()
        v.ExecuteNonQuery()
        Con.Close()
        v = Nothing
    End Sub

    Public Sub Delete_Buy(ByVal ID2 As Integer)
        Dim b As New OleDbCommand("Delete from The_Table where ID=@ID2", Con)
        b.Parameters.Add("ID2", OleDbType.Integer).Value = ID2
        Con.Open()
        b.ExecuteNonQuery()
        Con.Close()
        b = Nothing
    End Sub

End Module
