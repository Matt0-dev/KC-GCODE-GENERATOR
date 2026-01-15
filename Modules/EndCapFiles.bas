Attribute VB_Name = "Module4"
Sub GenerateCNCEndCapFiles()

    ' Generate CNC files for EndCaps using C+X+X+X+C equation pattern
    ' C = 11.75, X values      = 11.25, 22.75, 34.75, 46.75
    ' Length range: 40 to 120 inches
    
    On Error GoTo ErrorHandler
    
    ' Constants
    Const C_VALUE As Double = 12
    Const MIN_LENGTH As Double = 40
    Const MAX_LENGTH As Double = 120
    Const WIDTH_VALUE As Double = 11.75
    
    ' Panel sizes (X values)
    Dim panelSizes As Variant
    panelSizes = Array(12, 23, 35, 47) ' made edit regarding these panels before we were calculating at .25 and .75 now we round out and do only solid numbers\
    
    
    ' Create the main folder structure
    Dim mainFolderPath As String
    mainFolderPath = Environ("USERPROFILE") & "\OneDrive\Desktop\CNCendCap\"
     
    ' Delete existing folder if it exists
    On Error Resume Next
    Kill mainFolderPath & "Male\*.*"
    Kill mainFolderPath & "Female\*.*"
    RmDir mainFolderPath & "Male"
    RmDir mainFolderPath & "Female"
    Kill mainFolderPath & "*.*"
    RmDir mainFolderPath
    On Error GoTo 0
    
    ' Create new folder structure
    On Error Resume Next
    MkDir mainFolderPath
    MkDir mainFolderPath & "Male\"
    MkDir mainFolderPath & "Female\"
    On Error GoTo 0
    
    ' Set width values (B6 and B8)
    Range("B6").Value = WIDTH_VALUE
    Range("B8").Value = WIDTH_VALUE
    
    ' Generate all possible combinations
    Dim maxPanels As Integer
    Dim currentLength As Double
    Dim i As Integer, j As Integer
    Dim panelCombination() As Double
    Dim fileName As String
    Dim pocketPosition As Double
    
    ' Try different numbers of panels (1 to max possible)
    For maxPanels = 1 To 10 ' Reasonable upper limit
        
        ' Generate all combinations with this number of panels
        Call GenerateCombinations(panelSizes, maxPanels, C_VALUE, MIN_LENGTH, MAX_LENGTH, mainFolderPath)
        
    Next maxPanels
    
    MsgBox "EndCap files created successfully!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbExclamation
End Sub

Sub GenerateCombinations(panelSizes As Variant, numPanels As Integer, C_VALUE As Double, MIN_LENGTH As Double, MAX_LENGTH As Double, mainFolderPath As String)
    ' Generate all combinations with specified number of panels
    Dim indices() As Integer
    ReDim indices(1 To numPanels)
    Dim i As Integer
    
    ' Initialize indices
    For i = 1 To numPanels
        indices(i) = 0
    Next i
    
    ' Generate combinations
    Do
        ' Calculate current length
        Dim currentLength As Double
        currentLength = 2 * C_VALUE ' Start and end C values
         
        Dim panelConfig As String
        panelConfig = "C"
        
        For i = 1 To numPanels
            currentLength = currentLength + panelSizes(indices(i))
            panelConfig = panelConfig & "_" & Format(panelSizes(indices(i)), "0")
        Next i
        panelConfig = panelConfig & "_C"
        
        ' Check if length is within bounds
        If currentLength >= MIN_LENGTH And currentLength <= MAX_LENGTH Then
            ' Create files for this configuration
            Call CreateEndCapFile(panelConfig, currentLength, panelSizes, indices, numPanels, C_VALUE, mainFolderPath)
        End If
        
        ' Increment indices (like counting in base 4)
        Dim carry As Boolean
        carry = True
        For i = 1 To numPanels
            If carry Then
                indices(i) = indices(i) + 1
                If indices(i) > 3 Then
                    indices(i) = 0
                Else
                    carry = False
                End If
            End If
        Next i
        
        If carry Then Exit Do ' All combinations tried
        
    Loop
End Sub

Sub CreateEndCapFile(panelConfig As String, totalLength As Double, panelSizes As Variant, indices() As Integer, numPanels As Integer, C_VALUE As Double, mainFolderPath As String)
    ' Clear J6 to J13
    Range("J6:J13").ClearContents
    
    ' Set B7 (height/length)
    Range("B7").Value = totalLength
    
    ' Set pockets based on totalLength
    If totalLength > 84 Then
    Range("F7").Value = 10
    Range("F9").Formula = "=Height/2"
        Else
    Range("F7").Value = 20
    Range("F9").Value = 0
        End If
                
    ' Set pocket locations
    ' First C pocket at J6
    Range("J6").Value = 8
    
    ' Calculate pocket positions for each panel
    Dim currentPosition As Double
    currentPosition = C_VALUE ' Start after first C
    
    Dim pocketIndex As Integer
    pocketIndex = 7 ' Start at J7
    
    Dim i As Integer
    For i = 1 To numPanels
        Dim panelSize As Double
        panelSize = panelSizes(indices(i))
        
        ' Set pockets based on panel size
        Select Case panelSize
            Case 12 ' made edit to make into solid number
                ' Pocket at 4" from panel start
                If pocketIndex <= 13 Then
                    Range("J" & pocketIndex).Value = currentPosition + 4
                    pocketIndex = pocketIndex + 1
                End If
                
            Case 23 ' made edit to make into solid number
                ' Pocket at 10" from panel start
                If pocketIndex <= 13 Then
                    Range("J" & pocketIndex).Value = currentPosition + 10
                    pocketIndex = pocketIndex + 1
                End If
                
            Case 35 ' made edit to make into solid number
                ' Two pockets: at 10" and at (34.75-10)=24.75" from panel start
                If pocketIndex <= 13 Then
                    Range("J" & pocketIndex).Value = currentPosition + 10
                    pocketIndex = pocketIndex + 1
                End If
                If pocketIndex <= 13 Then
                    Range("J" & pocketIndex).Value = currentPosition + 25 ' made edit to make into solid number
                    pocketIndex = pocketIndex + 1
                End If
                
            Case 47 ' made edit to make into solid number
                ' Two pockets: at 10" and at (46.75-10)=36.75" from panel start
                If pocketIndex <= 13 Then
                    Range("J" & pocketIndex).Value = currentPosition + 10
                    pocketIndex = pocketIndex + 1
                End If
                If pocketIndex <= 13 Then
                    Range("J" & pocketIndex).Value = currentPosition + 37 ' made edit to make into solid number
                    pocketIndex = pocketIndex + 1
                End If
        End Select
        
        currentPosition = currentPosition + panelSize
    Next i
    
    ' Last C pocket at J13 (or next available cell)
    If pocketIndex <= 13 Then
        Range("J13").Value = totalLength - 8
    ElseIf pocketIndex = 14 Then
        ' If we've used all cells up to J12, put the last C pocket value in J13
        Range("J13").Value = totalLength - 8
    End If
    
    ' Create Male file
    Dim maleFileNumber As Integer
    maleFileNumber = FreeFile
    Dim maleFilePath As String
    maleFilePath = mainFolderPath & "Male\" & panelConfig & ".cnc"
    
    Open maleFilePath For Output As maleFileNumber
    Print #maleFileNumber, Range("C32").Value ' Male G-code from C32
    Close maleFileNumber
    
    ' Create Female file
    Dim femaleFileNumber As Integer
    femaleFileNumber = FreeFile
    Dim femaleFilePath As String
    femaleFilePath = mainFolderPath & "Female\" & panelConfig & ".cnc"
    
    Open femaleFilePath For Output As femaleFileNumber
    Print #femaleFileNumber, Range("C30").Value ' Female G-code from C30
    Close femaleFileNumber
    
End Sub
