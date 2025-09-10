Attribute VB_Name = "AutoUpdateAnalysis"
Option Explicit
 
Sub AutoUpdateAnalysis()
    ' This macro automatically updates all analysis sheets when data changes
    ' Call this from Workbook_Open or Worksheet_Change events
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range
    
    ' Check if we have data in DADOS_BRUTOS sheet
    Set ws = ThisWorkbook.Sheets("DADOS_BRUTOS")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow <= 1 Then
        MsgBox "No data found in DADOS_BRUTOS sheet", vbExclamation
        GoTo CleanUp
    End If
    
    ' Set data range (excluding headers)
    Set dataRange = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))
    
    ' Call Python script to update analysis (cross-platform)
    Call UpdateAnalysisWithPython
    
    ' Refresh all pivot tables if any
    Dim pt As PivotTable
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
    
    ' Update formulas
    ThisWorkbook.Calculate
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in AutoUpdateAnalysis: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

Sub UpdateAnalysisWithPython()
    ' Cross-platform call to Python update script
    ' Honors env var DIVERSITY_PYTHON if set to full path of python exe
    
    Dim pythonScript As String
    Dim excelPath As String
    Dim pyPath As String
    Dim cmd As String
    
    On Error GoTo PythonError
    
    excelPath = ThisWorkbook.FullName
    
    ' Script: procurar na mesma pasta ou em subpasta python_pipeline
    Dim cand1 As String, cand2 As String
    cand1 = ThisWorkbook.Path & Application.PathSeparator & "update_excel_analysis.py"
    cand2 = ThisWorkbook.Path & Application.PathSeparator & "python_pipeline" & Application.PathSeparator & "update_excel_analysis.py"
    If Dir(cand1) <> "" Then
        pythonScript = cand1
    ElseIf Dir(cand2) <> "" Then
        pythonScript = cand2
    Else
        MsgBox "Script não encontrado nas pastas esperadas: " & cand1 & " ou " & cand2, vbExclamation
        Exit Sub
    End If
    
    pyPath = Environ$("DIVERSITY_PYTHON")
    If Len(pyPath) = 0 Then
        If InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0 Then
            ' macOS defaults
            If Dir("/opt/anaconda3/bin/python") <> "" Then
                pyPath = "/opt/anaconda3/bin/python"
            ElseIf Dir("/usr/local/bin/python3") <> "" Then
                pyPath = "/usr/local/bin/python3"
            Else
                pyPath = "python3"
            End If
        Else
            ' Windows default
            pyPath = "python"
        End If
    End If
    
    If InStr(1, Application.OperatingSystem, "Mac", vbTextCompare) > 0 Then
        ' macOS: use /bin/bash -lc to honor shell envs if any
        cmd = "/bin/bash -lc '" & Replace(pyPath, "'", "'\''") & " '" & Replace(pythonScript, "'", "'\''") & "' '" & Replace(excelPath, "'", "'\''") & "''"
        Shell cmd, vbHide
    Else
        ' Windows: WScript.Shell with wait
        Dim wsh As Object
        Set wsh = CreateObject("WScript.Shell")
        cmd = """" & pyPath & """ """ & pythonScript & """ """ & excelPath & """"
        wsh.Run cmd, 0, True
    End If
    
    Exit Sub
PythonError:
    MsgBox "Erro ao executar Python: " & Err.Description & Chr(10) & _
           "Defina DIVERSITY_PYTHON com o caminho do Python ou ajuste a macro.", vbCritical
End Sub

Sub AddAutoUpdateButton()
    ' This adds a button to the sheet to manually trigger updates
    
    Dim ws As Worksheet
    Dim btn As Button
    
    ' Add button to first sheet
    Set ws = ThisWorkbook.Sheets(1)
    
    ' Remove existing button if any
    On Error Resume Next
    ws.Buttons("AutoUpdateBtn").Delete
    On Error GoTo 0
    
    ' Create new button
    Set btn = ws.Buttons.Add(100, 10, 150, 30)
    With btn
        .Name = "AutoUpdateBtn"
        .Caption = "Atualizar Análises (Python)"
        .OnAction = "AutoUpdateAnalysis"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    MsgBox "Auto-update button added successfully!", vbInformation
End Sub

Sub Workbook_Open()
    ' This runs automatically when workbook opens
    
    On Error Resume Next
    Call AddAutoUpdateButton
    On Error GoTo 0
    
    ' Opcional: lembrar o usuário
    MsgBox "Clique em 'Atualizar Análises (Python)' para regerar as abas após editar DADOS_BRUTOS." & vbCrLf & _
           "Dica: defina a variável de ambiente DIVERSITY_PYTHON para o caminho do Python, se necessário.", vbInformation
End Sub
