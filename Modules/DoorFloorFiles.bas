Attribute VB_Name = "Module6"
Sub GenerateCNCDoorFloorFiles()

    ' Generate Door Floor CNC Files
    ' Fixed width: 46.75
    ' Uses C28 for G-code
    ' J15 = 10, J16 = 0, J17 = 10
    
    On Error GoTo ErrorHandler

    ' Clear cells J15-J17 and set values
    Range("J15:J17").ClearContents
    Range("J15").Value = 10
    Range("J16").Value = 0  ' Set J16 to 0
    Range("J17").Value = 10

    ' Set values for B7
    Dim valueB7 As Double

    ' Create the main CNCDoorFloor folder on the desktop
    Dim mainFolderPath As String
    mainFolderPath = Environ("USERPROFILE") & "\OneDrive\Desktop\CNCDoorFloor\"

    ' Delete the CNCDoorFloor folder if it exists
    On Error Resume Next
    RmDir mainFolderPath & "DoorFloor"
    Kill mainFolderPath & "*.*"
    RmDir mainFolderPath
    On Error GoTo 0

    ' Create the main folder if it doesn't exist
    On Error Resume Next
    MkDir mainFolderPath
    On Error GoTo 0

    ' Create the DoorFloor folder
    MkDir mainFolderPath & "DoorFloor\"

    ' Fixed width value
    Const DOOR_FLOOR_WIDTH As Double = 46.75

    Dim i As Double
    For i = 60 To 128 Step 0.25
        ' Create folders in DoorFloor
        MkDir mainFolderPath & "DoorFloor\" & Format(i, "0.0") & "-Inch\"

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

        ' Process DoorFloor
        ' Set B6 and B8 to the door floor width
        Range("B6").Value = DOOR_FLOOR_WIDTH
        Range("B8").Value = DOOR_FLOOR_WIDTH
        
        ' Reset F15, F16, F17 and J cell values
        Range("F15").Value = 0  ' Reset F15 to 0
        Range("F16").Value = 0  ' Reset F16 to 0
        Range("F17").Value = 0  ' Reset F17 to 0
        Range("J15:J17").ClearContents  ' Clear J15-J17

        ' Set J cell values (same as last program)
        Range("J15").Value = 10
        Range("J16").Value = 0
        Range("J17").Value = 10

        ' Create file path and name for DoorFloor
        Dim doorFloorFileNumber As Integer
        doorFloorFileNumber = FreeFile
        Dim doorFloorFilePath As String
        doorFloorFilePath = mainFolderPath & "DoorFloor\" & Format(i, "0.0") & "-Inch\" & Format(DOOR_FLOOR_WIDTH, "0.0") & "x" & Format(i, "0.0") & ".cnc"

        ' Open the file for writing
        Open doorFloorFilePath For Output As doorFloorFileNumber

        ' Copy and paste C28 value to the file
        Print #doorFloorFileNumber, Range("C28").Value

        ' Close the file
        Close doorFloorFileNumber

    Next i

    MsgBox "Door Floor files created successfully!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbExclamation
End Sub
