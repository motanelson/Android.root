Attribute VB_Name = "Módulo3"
Option Compare Database
Option Explicit

Sub ImportarCSV_ComValidacao()

    On Error GoTo ERRO

    Dim db As DAO.Database
    Dim f As Integer, ferr As Integer
    Dim linha As String
    Dim campos() As String
    Dim sql As String

    Dim ficheiroCSV As String
    Dim tabela As String
    Dim numCampos As Integer
    Dim sep As String

    ' === Perguntas ao utilizador ===
    ficheiroCSV = EscolherCSV()
    If ficheiroCSV = "" Then Exit Sub

    tabela = InputBox("Nome da tabela destino:")
    If tabela = "" Then Exit Sub

    numCampos = Val(InputBox("Número de campos por linha:"))
    If numCampos <= 0 Then Exit Sub

    sep = InputBox("Separador (ex: , ; | )", ",")
    If sep = "" Then sep = ","

    Set db = CurrentDb

    ' === Abrir ficheiros ===
    f = FreeFile
    Open ficheiroCSV For Input As #f

    ferr = FreeFile
    Open CurrentProject.Path & "\error.csv" For Output As #ferr

    ' === Ler CSV linha a linha ===
    Do While Not EOF(f)

        Line Input #f, linha
        campos = Split(linha, sep)

        ' Verificar nº de colunas
        If UBound(campos) + 1 <> numCampos Then
            Print #ferr, linha
        Else
            ' Construir INSERT
            sql = "INSERT INTO [" & tabela & "] VALUES ("

            Dim i As Integer
            For i = 0 To numCampos - 1
                sql = sql & "'" & Replace(campos(i), "'", "''") & "'"
                If i < numCampos - 1 Then sql = sql & ","
            Next

            sql = sql & ")"
            'MsgBox sql
            db.Execute sql, dbFailOnError
        End If

    Loop

    Close #f
    Close #ferr

    MsgBox "Importação concluída!", vbInformation
    Exit Sub

ERRO:
    Close
    MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Function EscolherCSV() As String
    Dim fd As Object

    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker

    With fd
        .Title = "Escolher ficheiro CSV"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .AllowMultiSelect = False

        If .Show = -1 Then
            EscolherCSV = .SelectedItems(1)
        Else
            EscolherCSV = ""
        End If
    End With
End Function

