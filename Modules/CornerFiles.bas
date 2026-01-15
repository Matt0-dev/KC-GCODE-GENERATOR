Attribute VB_Name = "Module2"
Sub GenerateCNCCornerFiles()

    ' Generate Corner CNC Files
    ' Fixed width: 19.5
    ' F15: 4, F17: 15.5
    
    On Error GoTo ErrorHandler

    ' Clear cells J15-J17 and set values
    Range("J15:J17").ClearContents
    Range("J15").Value = 10
    Range("J16").Value = 0  ' Set J16 to 0
    Range("J17").Value = 10

    ' Set values for B7
    Dim valueB7 As Double

    ' Create the main CNCCorner folder on the desktop
    Dim mainFolderPath As String
    mainFolderPath = Environ("USERPROFILE") & "\OneDrive\Desktop\CNCCorner\"

    ' Delete the CNCCorner folder if it exists
    On Error Resume Next
    RmDir mainFolderPath & "CornerFreezer"
    RmDir mainFolderPath & "CornerRefrigerator"
    Kill mainFolderPath & "*.*"
    RmDir mainFolderPath
    On Error GoTo 0

    ' Create the main folder if it doesn't exist
    On Error Resume Next
    MkDir mainFolderPath
    On Error GoTo 0

    ' Create the CornerFreezer and CornerRefrigerator folders
    MkDir mainFolderPath & "CornerFreezer\"
    MkDir mainFolderPath & "CornerRefrigerator\"

    ' Fixed width value
    Const CORNER_WIDTH As Double = 19.5
    
    ' Fixed pocket positions
    Const F15_VALUE As Double = 4
    Const F17_VALUE As Double = 15.5

    Dim i As Double
    For i = 60 To 128 Step 0.25
        ' Create folders in CornerFreezer and CornerRefrigerator
        MkDir mainFolderPath & "CornerFreezer\" & Format(i, "0.0") & "-Inch\"
        MkDir mainFolderPath & "CornerRefrigerator\" & Format(i, "0.0") & "-Inch\"

        ' Set value for B7 based on the current folder
        Range("B7").Value = i

        ' Set F6 and F9 values based on B7
        If i < 80.5 Then
            Range("F7").Value = 10
            Range("F9").Value = 0
        Else
            Range("F7").Value = 10
            Range("F9").Formula = "=Height/2"
        End If

        ' Process CornerFreezer
        ' Set B6 and B8 to the corner width
        Range("B6").Value = CORNER_WIDTH
        Range("B8").Value = CORNER_WIDTH
        
        ' Reset F15, F16, F17 and J cell values
        Range("F15").Value = 0  ' Reset F15 to 0
        Range("F16").Value = 0  ' Reset F16 to 0
        Range("F17").Value = 0  ' Reset F17 to 0
        Range("J15:J17").ClearContents  ' Clear J15-J17

        ' Set F15 and F17 values for corner
        Range("F15").Value = F15_VALUE
        Range("F17").Value = F17_VALUE

        ' Set J cell values
        Range("J15").Value = 10
        Range("J16").Value = 0
        Range("J17").Value = 10

        ' Create file path and name for CornerFreezer
        Dim cornerFreezerFileNumber As Integer
        cornerFreezerFileNumber = FreeFile
        Dim cornerFreezerFilePath As String
        cornerFreezerFilePath = mainFolderPath & "CornerFreezer\" & Format(i, "0.0") & "-Inch\" & Format(CORNER_WIDTH, "0.0") & "x" & Format(i, "0.0") & ".cnc"

        ' Open the file for writing
        Open cornerFreezerFilePath For Output As cornerFreezerFileNumber

        ' Copy and paste C22 value to the file (same as regular freezer)
        Print #cornerFreezerFileNumber, Range("C22").Value

        ' Close the file
        Close cornerFreezerFileNumber

        ' Process CornerRefrigerator
        ' Set B6 and B8 to the corner width
        Range("B6").Value = CORNER_WIDTH
        Range("B8").Value = CORNER_WIDTH
        
        ' Reset F15, F16, F17 and J cell values
        Range("F15").Value = 0  ' Reset F15 to 0
        Range("F16").Value = 0  ' Reset F16 to 0
        Range("F17").Value = 0  ' Reset F17 to 0
        Range("J15:J17").ClearContents  ' Clear J15-J17

        ' Set F15 and F17 values for corner
        Range("F15").Value = F15_VALUE
        Range("F17").Value = F17_VALUE

        ' Set J cell values
        Range("J15").Value = 10
        Range("J16").Value = 0
        Range("J17").Value = 10

        ' Create file path and name for CornerRefrigerator
        Dim cornerRefrigeratorFileNumber As Integer
        cornerRefrigeratorFileNumber = FreeFile
        Dim cornerRefrigeratorFilePath As String
        cornerRefrigeratorFilePath = mainFolderPath & "CornerRefrigerator\" & Format(i, "0.0") & "-Inch\" & Format(CORNER_WIDTH, "0.0") & "x" & Format(i, "0.0") & ".cnc"

        ' Open the file for writing
        Open cornerRefrigeratorFilePath For Output As cornerRefrigeratorFileNumber

        ' Copy and paste C24 value to the file (same as regular refrigerator)
        Print #cornerRefrigeratorFileNumber, Range("C24").Value

        ' Close the file
        Close cornerRefrigeratorFileNumber

    Next i

    MsgBox "Corner files created successfully!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbExclamation
End Sub
