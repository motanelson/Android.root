Attribute VB_Name = "Módulo2"
Option Compare Database

Sub f1()
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim aaa As String
    aaa = InputBox("give the sql")
    db.Execute aaa, dbFailOnError

    Set db = Nothing
End Sub

