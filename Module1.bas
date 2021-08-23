Attribute VB_Name = "Module1"
Sub openbtn1()

Dim v As String
Dim Shex As Object
   'Range("H20").Select
   Set Shex = CreateObject("Shell.Application")
   v = Range("H6").Value
   tgtfile = v
   Shex.Open (tgtfile)
   'Shex.Open (v)
End Sub

Sub openbtn2()

Dim v2 As String
Dim Shex As Object
   'Range("H21").Select
   Set Shex = CreateObject("Shell.Application")
   v2 = Range("H7").Value
   tgtfile = v2
   Shex.Open (tgtfile)
   'Shex.Open (v)
End Sub
Sub openbtn3()

Dim v3 As String
Dim Shex As Object
   'Range("H22").Select
   Set Shex = CreateObject("Shell.Application")
   v3 = Range("H8").Value
   tgtfile = v3
   Shex.Open (tgtfile)
   'Shex.Open (v)
End Sub

Sub openbtn4()

Dim v4 As String
Dim Shex As Object
   'Range("H23").Select
   Set Shex = CreateObject("Shell.Application")
   v4 = Range("H9").Value
   tgtfile = v4
   Shex.Open (tgtfile)
   'Shex.Open (v4)
End Sub
Sub openbtnall4()

Dim v1 As String
Dim v2 As String
Dim v3 As String
Dim v4 As String
Dim Shex As Object
   
   'Range("H20").Select
   v1 = Range("H6").Value
   'Range("H21").Select
   v2 = Range("H7").Value
   'Range("H22").Select
   v3 = Range("H8").Value
   'Range("H23").Select
   v4 = Range("H9").Value

   
   Set Shex = CreateObject("Shell.Application")
   tgtfile = v1
   Shex.Open (tgtfile)
   
   tgtfile = v2
   Shex.Open (tgtfile)
   
   tgtfile = v3
   Shex.Open (tgtfile)
   
   tgtfile = v4
   Shex.Open (tgtfile)

End Sub

Sub openbtnall3()

Dim v1 As String
Dim v2 As String
Dim v3 As String

Dim Shex As Object
   
   'Range("H21").Select
   v1 = Range("H7").Value
   'Range("H22").Select
   v2 = Range("H8").Value
   'Range("H23").Select
   v3 = Range("H9").Value

   
   Set Shex = CreateObject("Shell.Application")
   tgtfile = v1
   Shex.Open (tgtfile)
   
   tgtfile = v2
   Shex.Open (tgtfile)
   
   tgtfile = v3
   Shex.Open (tgtfile)
   
End Sub

Sub godowncopy()
' function select next right cell and copy

    ActiveCell.Offset(1, 0).Select ' ActiveCell.Offset(collum, line).Select
    
    Selection.Copy


End Sub
Sub gorightcopy()
' function select next right cell and copy

    ActiveCell.Offset(0, 1).Select ' ActiveCell.Offset(collum, line).Select
    
    Selection.Copy


End Sub
Sub FnGetSheetsName()
' get all worksheet names and paste in actual cell line above line

Dim mainworkBook As Workbook

Set mainworkBook = ActiveWorkbook

For i = 1 To mainworkBook.Sheets.Count

ActiveCell.Value = mainworkBook.Sheets(i).Name
ActiveCell.Offset(1, 0).Select

Next i

End Sub
Sub getfilelistfromfolder()

Dim varDirectory As Variant
Dim flag As Boolean
Dim i As Integer
Dim strDirectory As String
Dim myextension As String

strDirectory = Application.ActiveWorkbook.Path & "\"
i = 1
flag = True

' select and delet all data in A column
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents
Range("A2").Select

'MsgBox (strDirectory)
'varDirectory = Dir("C:\Macro\*.dwg", vbNormal)
myextension = InputBox("Type files extension, ex: dwg, pdf, txt")
varDirectory = Dir(strDirectory & "*" & myextension, vbNormal)

While flag = True
    If varDirectory = "" Then
        flag = False
    Else
        'Cells(i + 1, 1) = varDirectory
        Cells(i + 1, 1) = varDirectory '1 column A, 2 column B and go on
        varDirectory = Dir
        i = i + 1
    End If
Wend
End Sub
Sub RenameFiles()

Dim xDir As String
Dim xFile As String
Dim xRow As Long

    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
    
        If .Show = -1 Then
            xDir = .SelectedItems(1)
            xFile = Dir(xDir & Application.PathSeparator & "*")
            
            Do Until xFile = ""
                xRow = 0
                On Error Resume Next
                xRow = Application.Match(xFile, Range("A:A"), 0)
                If xRow > 0 Then
                    Name xDir & Application.PathSeparator & xFile As _
                    xDir & Application.PathSeparator & Cells(xRow, "E").Value
                End If
                xFile = Dir
            Loop
        End If
    End With
End Sub




