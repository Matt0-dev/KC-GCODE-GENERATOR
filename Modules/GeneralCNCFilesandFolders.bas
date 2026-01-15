Attribute VB_Name = "Module5"
Sub GenerateCNCFilesAndFolders()

    ' EDITED 11/10/25 - Added MaleConnectedCeiling and FemaleConnectedCeiling
    

    On Error GoTo ErrorHandler

    ' Clear cells J15-J17 and set values
    Range("J15:J17").ClearContents
    Range("J15").Value = 10
    Range("J16").Value = 0  ' Set J16 to 0
    Range("J17").Value = 10

    ' Set values for B7
    Dim valueB7 As Double

' Create the main CNCTEST folder on the desktop
Dim mainFolderPath As String
mainFolderPath = Environ("USERPROFILE") & "\OneDrive\Desktop\CNCTEST75\"

' Delete the CNCTEST folder if it exists
On Error Resume Next
RmDir mainFolderPath & "Freezer"
RmDir mainFolderPath & "Refrigerator"
RmDir mainFolderPath & "Ceiling"
RmDir mainFolderPath & "MaleConnectedCeiling"
RmDir mainFolderPath & "FemaleConnectedCeiling"
Kill mainFolderPath & "*.*"
RmDir mainFolderPath
On Error GoTo 0

' Create the main folder if it doesn't exist
On Error Resume Next
MkDir mainFolderPath
On Error GoTo 0

' Create the Freezer, Refrigerator, Ceiling, MaleConnectedCeiling, and FemaleConnectedCeiling folders
MkDir mainFolderPath & "Freezer\"
MkDir mainFolderPath & "Refrigerator\"
MkDir mainFolderPath & "Ceiling\"
MkDir mainFolderPath & "MaleConnectedCeiling\"
MkDir mainFolderPath & "FemaleConnectedCeiling\"

Dim i As Double
    For i = 60 To 128 Step 0.25
        ' Create folders in Freezer, Refrigerator, Ceiling, MaleConnectedCeiling, and FemaleConnectedCeiling
        MkDir mainFolderPath & "Freezer\" & Format(i, "0.0") & "-Inch\"
        MkDir mainFolderPath & "Refrigerator\" & Format(i, "0.0") & "-Inch\"
        MkDir mainFolderPath & "Ceiling\" & Format(i, "0.0") & "-Inch\"
        MkDir mainFolderPath & "MaleConnectedCeiling\" & Format(i, "0.0") & "-Inch\"
        MkDir mainFolderPath & "FemaleConnectedCeiling\" & Format(i, "0.0") & "-Inch\"

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

        ' Loop through B6 values (46.75, 34.75, and 22.75)
        Dim b6Values As Variant
        b6Values = Array(46.75, 34.75, 22.75)

        ' Loop through folders in Freezer and create files
        For Each b6File In b6Values
            ' Set B6 and B8 to the same value
            Range("B6").Value = b6File
            Range("B8").Value = b6File
            
            ' Reset F15, F16, F17 and J cell values for each iteration
            Range("F15").Value = 0  ' Reset F15 to 0
            Range("F16").Value = 0  ' Reset F16 to 0
            Range("F17").Value = 0  ' Reset F17 to 0
            Range("J15:J17").ClearContents  ' Clear J15-J17

            ' Set F15 and F17 values based on B6 for Freezer
            If b6File = 46.75 Or b6File = 34.75 Then
                Range("F15").Value = 10
                Range("F17").Formula = "=IF(OR(B6<22.9,HPocket_X0=0),0,Width-HPocket_X0)"
            ElseIf b6File = 22.75 Then
                Range("F15").Value = 10
                Range("F17").Value = 0
            End If

            ' Set J cell values based on B6
            If b6File = 46.75 Or b6File = 34.75 Then
                Range("J15").Value = 10
                Range("J16").Value = 0
                Range("J17").Value = 10
            ElseIf b6File = 22.5 Then
                Range("J15").Value = 10
                Range("J16").Value = 0
                Range("J17").Value = 0
                
                
                
                
            End If

            ' Create file path and name for Freezer
            Dim freezerFileNumber As Integer
            freezerFileNumber = FreeFile
            Dim freezerFilePath As String
            freezerFilePath = mainFolderPath & "Freezer\" & Format(i, "0.0") & "-Inch\" & Format(b6File, "0.0") & "x" & Format(i, "0.0") & ".cnc"

            ' Open the file for writing
            Open freezerFilePath For Output As freezerFileNumber

            ' Copy and paste C22 value to the file
            Print #freezerFileNumber, Range("C22").Value

            ' Close the file
            Close freezerFileNumber
        Next b6File

        ' Repeat the process for Refrigerator
        For Each b6File In b6Values
            ' Set B6 and B8 to the same value
            Range("B6").Value = b6File
            Range("B8").Value = b6File
            
            ' Reset F15 and F17 values explicitly for each iteration
            Range("F15").Value = 0  ' Reset F15 to 0
            Range("F17").Value = 0  ' Reset F17 to 0
            Range("J15:J17").ClearContents  ' Clear J15-J17

            ' Set F15 and F17 values based on B6 for Refrigerator
            If b6File = 46.75 Or b6File = 34.75 Then
                Range("F15").Value = 10
                Range("F17").Formula = "=IF(OR(B6<22.9,HPocket_X0=0),0,Width-HPocket_X0)"
            ElseIf b6File = 22.75 Then
                Range("F15").Value = 10
                Range("F17").Value = 0
            End If

            ' Set J cell values based on B6 for Ceiling
            If b6File = 46.75 Or b6File = 34.75 Then
                Range("J15").Value = 10
                Range("J16").Value = 0
                Range("J17").Value = 10
            ElseIf b6File = 22.75 Then
                Range("J15").Value = 10
                Range("J16").Value = 0
                Range("J17").Value = 0
                
            End If

            ' Create file path and name for Refrigerator
            Dim refrigeratorFileNumber As Integer
            refrigeratorFileNumber = FreeFile
            Dim refrigeratorFilePath As String
            refrigeratorFilePath = mainFolderPath & "Refrigerator\" & Format(i, "0.0") & "-Inch\" & Format(b6File, "0.0") & "x" & Format(i, "0.0") & ".cnc"

            ' Open the file for writing
            Open refrigeratorFilePath For Output As refrigeratorFileNumber

            ' Copy and paste C24 value to the file
            Print #refrigeratorFileNumber, Range("C24").Value

            ' Close the file
            Close refrigeratorFileNumber
        Next b6File

        ' Repeat the process for Ceiling
        For Each b6File In b6Values
            ' Set B6 and B8 to the same value
            Range("B6").Value = b6File
            Range("B8").Value = b6File
            
            ' Reset F15 and J cell values for each iteration
            Range("F15").Value = 0  ' Reset F15 to 0
            Range("J15:J17").ClearContents  ' Clear J15-J17

            ' Set J cell values based on B6 for Ceiling
            If b6File = 46.75 Or b6File = 34.75 Then
                Range("J15").Value = 10
                Range("J16").Value = 0
                Range("J17").Value = 10
            ElseIf b6File = 22.75 Then
                Range("J15").Value = 10
                Range("J16").Value = 0
                Range("J17").Value = 0
                
            End If

            ' Set value for B7 based on the current folder for Ceiling
            If i > 84 Then
                Range("F7").Value = 10
                Range("F9").Formula = "=Height/2"
            Else
                Range("F7").Value = 20
                Range("F9").Value = 0
            End If

            ' Create file path and name for Ceiling
            Dim ceilingFileNumber As Integer
            ceilingFileNumber = FreeFile
            Dim ceilingFilePath As String
            ceilingFilePath = mainFolderPath & "Ceiling\" & Format(i, "0.0") & "-Inch\" & Format(b6File, "0.0") & "x" & Format(i, "0.0") & ".cnc"

            ' Open the file for writing
            Open ceilingFilePath For Output As ceilingFileNumber

            ' Copy and paste C26 value to the file
            Print #ceilingFileNumber, Range("C26").Value

            ' Close the file
            Close ceilingFileNumber
        Next b6File
        
        ' Process MaleConnectedCeiling (ceiling with wall-style F values)
        For Each b6File In b6Values
            ' Set B6 and B8 to the same value
            Range("B6").Value = b6File
            Range("B8").Value = b6File
            
            ' Reset F15 and J cell values for each iteration
            Range("F15").Value = 0  ' Reset F15 to 0
            Range("J15:J17").ClearContents  ' Clear J15-J17

            ' Set F15 and F17 values based on B6 (like walls)
            If b6File = 46.75 Or b6File = 34.75 Then
                Range("F15").Value = 10
                Range("F17").Value = "=IF(OR(B6<22.9,HPocket_X0=0),0,Width-HPocket_X0)"
            ElseIf b6File = 22.75 Then
                Range("F15").Value = 10
                Range("F17").Value = 0
            End If

            ' Set J cell values based on B6 for MaleConnectedCeiling
            If b6File = 46.75 Or b6File = 34.75 Then
                Range("J15").Value = 10
                Range("J16").Value = 0
                Range("J17").Value = 10
            ElseIf b6File = 22.75 Then
                Range("J15").Value = 10
                Range("J16").Value = 0
                Range("J17").Value = 0
                
            End If

            ' Set value for B7 based on the current folder for MaleConnectedCeiling
            If i > 70 Then
                Range("F7").Value = 10
                Range("F9").Formula = "=Height/2"
            Else
                Range("F7").Value = 20
                Range("F9").Value = 0
            End If

            ' Create file path and name for MaleConnectedCeiling
            Dim maleConnectedCeilingFileNumber As Integer
            maleConnectedCeilingFileNumber = FreeFile
            Dim maleConnectedCeilingFilePath As String
            maleConnectedCeilingFilePath = mainFolderPath & "MaleConnectedCeiling\" & Format(i, "0.0") & "-Inch\" & Format(b6File, "0.0") & "x" & Format(i, "0.0") & ".cnc"

            ' Open the file for writing
            Open maleConnectedCeilingFilePath For Output As maleConnectedCeilingFileNumber

            ' Copy and paste C36 value to the file
            Print #maleConnectedCeilingFileNumber, Range("C36").Value

            ' Close the file
            Close maleConnectedCeilingFileNumber
        Next b6File
        
        ' Process FemaleConnectedCeiling (ceiling with wall-style F values)
        For Each b6File In b6Values
            ' Set B6 and B8 to the same value
            Range("B6").Value = b6File
            Range("B8").Value = b6File
            
            ' Reset F15 and J cell values for each iteration
            Range("F15").Value = 0  ' Reset F15 to 0
            Range("J15:J17").ClearContents  ' Clear J15-J17

            ' Set F15 and F17 values based on B6 (like walls)
            If b6File = 46.75 Or b6File = 34.75 Then
                Range("F15").Value = 10
                Range("F17").Value = "=IF(OR(B6<22.9,HPocket_X0=0),0,Width-HPocket_X0)"
            ElseIf b6File = 22.75 Then
                Range("F15").Value = 10
                Range("F17").Value = 0
            End If

            ' Set J cell values based on B6 for FemaleConnectedCeiling
            If b6File = 46.75 Or b6File = 34.75 Then
                Range("J15").Value = 10
                Range("J16").Value = 0
                Range("J17").Value = 10
            ElseIf b6File = 22.75 Then
                Range("J15").Value = 10
                Range("J16").Value = 0
                Range("J17").Value = 0
                
            End If

            ' Set value for B7 based on the current folder for FemaleConnectedCeiling
            If i > 70 Then
                Range("F7").Value = 10
                Range("F9").Formula = "=Height/2"
            Else
                Range("F7").Value = 20
                Range("F9").Value = 0
            End If

            ' Create file path and name for FemaleConnectedCeiling
            Dim femaleConnectedCeilingFileNumber As Integer
            femaleConnectedCeilingFileNumber = FreeFile
            Dim femaleConnectedCeilingFilePath As String
            femaleConnectedCeilingFilePath = mainFolderPath & "FemaleConnectedCeiling\" & Format(i, "0.0") & "-Inch\" & Format(b6File, "0.0") & "x" & Format(i, "0.0") & ".cnc"

            ' Open the file for writing
            Open femaleConnectedCeilingFilePath For Output As femaleConnectedCeilingFileNumber

            ' Copy and paste C38 value to the file
            Print #femaleConnectedCeilingFileNumber, Range("C38").Value

            ' Close the file
            Close femaleConnectedCeilingFileNumber
        Next b6File
        
    Next i

    MsgBox "Folder structure and files created successfully!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbExclamation
End Sub
