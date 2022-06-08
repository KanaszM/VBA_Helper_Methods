Attribute VB_Name = "Helper_Methods"
Public Sub Log(sText As String, Optional bClear As Boolean = False)
   If bClear = True Then
      Application.SendKeys "^g^{END}", True
      DoEvents
      Debug.Print String(30, vbCrLf)
   End If
   Debug.Print sText
End Sub

Public Sub OptimizeVBA(bMode As Boolean)
    Dim oWS As Worksheet
    With Application
        If bMode Then
            .Calculation = xlCalculationManual
            .ScreenUpdating = False
            .EnableEvents = False
            .EnableAnimations = False
            .DisplayAlerts = False
            .PrintCommunication = False
            For Each oWS In ActiveWorkbook.Worksheets
                oWS.DisplayPageBreaks = False
            Next oWS
        Else
            .Calculation = xlCalculationAutomatic
            .ScreenUpdating = True
            .EnableEvents = True
            .EnableAnimations = True
            .DisplayAlerts = True
      .PrintCommunication = True
        End If
    End With
End Sub

Private Function CountFilesInFolder(sDIR As String, Optional sType As String)
    Dim oFile As Variant
    Dim iCount As Integer
    If Right(sDIR, 1) <> "\" Then sDIR = sDIR & "\"
    oFile = Dir(sDIR & sType)
    While (oFile <> "")
        iCount = iCount + 1
        oFile = Dir
    Wend
    CountFilesInFolder = iCount
End Function
