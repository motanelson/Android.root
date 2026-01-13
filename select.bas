Attribute VB_Name = "Módulo3"
Option Compare Database

Sub f1()
    On Error GoTo ERRO

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim i As Integer
    Dim linha As String

    Set db = CurrentDb
    sql = Trim(InputBox("Give the SQL command"))

    If sql = "" Then Exit Sub

    ' Se for SELECT → mostrar resultados
    If UCase(Left(sql, 6)) = "SELECT" Then

        Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

        If rs.EOF Then
            MsgBox "Sem resultados", vbInformation
            Exit Sub
        End If

        ' Cabeçalho
        linha = ""
        For i = 0 To rs.Fields.Count - 1
            linha = linha & rs.Fields(i).Name & vbTab
        Next
        linha = linha & vbCrLf & String(50, "-") & vbCrLf

        ' Dados
        Do While Not rs.EOF
            For i = 0 To rs.Fields.Count - 1
                linha = linha & Nz(rs.Fields(i).Value, "") & vbTab
            Next
            linha = linha & vbCrLf
            rs.MoveNext
        Loop

        MsgBox linha, vbInformation, "SQL Result"

        rs.Close
        Set rs = Nothing

    Else
        ' Comandos sem resultado
        db.Execute sql, dbFailOnError
        MsgBox "SQL executado com sucesso", vbInformation
    End If

    Exit Sub

ERRO:
    MsgBox "Erro:" & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
End Sub
