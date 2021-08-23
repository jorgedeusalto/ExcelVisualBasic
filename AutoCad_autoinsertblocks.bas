Attribute VB_Name = "Module1"

Sub createscripttocad()
'
' get circuit numbers and arrrange to make a script to autocad
'

' replace " A"(space A) just for "A" with no space before A

    Range("B2").Select
    Selection.NumberFormat = "@"
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:=" A", Replacement:="A", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("B1").Select

'put column D in capital letters
Dim range1 As Range
    For Each range1 In Range("B1:D200")
        range1.Offset(0, 0).Value = UCase(range1.Value)
    Next range1


'insert empty column B if not already has
    If Range("B1").Value <> "" Then
        Columns("B").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    End If

'' clear fake blank cells
'    Range("A1:C200").Select
'    With Selection.Interior
'        .Pattern = xlNone
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
'    End With
'
'    Selection.Replace What:="", Replacement:="x-x-x", LookAt:=xlPart, _
'        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
'        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
'    Selection.Replace What:="x-x-x", Replacement:="", LookAt:=xlPart, _
'        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
'        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

Range("A1").Select

'**********************change circuits names from L1/1 to 1L1, or L1/L2/L3/8  to 8L1 8L2 8L3

Dim CircuitNumber, current, devicerange As Range
Dim mcb3plight() As String ' split string when mcb 3p with lighting is found
Dim tx1, tx2, tx3 As String
Dim islight As String ' return string load name with or without lighting word
Dim rating As String 'device rating as 10A 20A 16A...
Dim type1 As String ' devide type like  'C' 'B'....

Dim f1 As Integer ' f1 return a position number (from right to left) when found "/" in a string
Dim Enumber As Integer ' number of lights emergencies circuits E1 E2...
Dim Enumbercount As Integer ' count 4 by 4 number of lights emergencies circuits e1 e1 e1 e1, e2 e2 e2 e2, e3 e3 e3 e3, ...
Dim CNLenght As Integer ' number of characters in circuit number, L1/1 = 4, L1/99=5, L1/L2/L3/1=10

Enumber = 1
Enumbercount = 1
 
    For Each CircuitNumber In Range("a1:a200")
    f1 = InStrRev(CircuitNumber.Value, "/") ' search for "/" from right to left and get the position
    islight = LCase(CircuitNumber.Columns(6).Value) 'in lower case get the string on device load description
    
        If f1 > 0 Then
                CNLenght = Len(CircuitNumber.Value) ' number of characters in circuit number, L1/1 = 4, L1/99=5, L1/L2/L3/1=10
                tx1 = Left(CircuitNumber.Value, f1 - 1) ' first part of the string circuitnumber 1 in L1/"1"
                tx2 = Right(CircuitNumber.Value, CNLenght - f1) ' rest of the string circuitnumber L1 in "L1"/1
                tx3 = tx2 & tx1 ' tx1 2 and tx1 together, it means 1L1, 1L2 ...
                rating = CircuitNumber.Columns(3).Value ' device rating like 10A, 20A, 36A...
                type1 = "'" & UCase(CircuitNumber.Columns(5).Value) & "'" ' device type like 'C' 'B'...
            
        '*** MCB's SP, mcbs single pole
            If CNLenght <= 6 And UCase(CircuitNumber.Columns(4).Value) = "MCB" And InStr(islight, "lighting") = 0 Then ' if total number of characters are iqual or less then 6, L1/1 turn in to 1L1
                CircuitNumber.Offset(0, 1).Value = tx3 & " " & rating & " " & type1 ' write on b column
            
        '*** MCB's SP, mcbs single pole **  WITH EMergency Lights  ** circuits
        '**********************************************************************
            ElseIf CNLenght <= 6 And UCase(CircuitNumber.Columns(4).Value) = "MCB" And InStr(islight, "lighting") <> 0 Then ' if total number of characters are iqual or less then 6, L1/1 turn in to 1L1
                CircuitNumber.Offset(0, 1).Value = "E" & Enumber & " " & tx3 & " " & rating & " " & type1
                    Enumbercount = Enumbercount + 1
                    If Enumbercount > 4 Then
                        Enumber = Enumber + 1
                        Enumbercount = 1
                    End If
                        
        '*** MCB's TP, mcbs triple pole
            ElseIf CNLenght > 6 And UCase(CircuitNumber.Columns(4).Value) = "MCB" And InStr(islight, "lighting") = 0 Then ' if total number of characters are bigger then 6, L1/L2/L3/8  to 8L1 8L2 8L3
                tx3 = Replace(tx3, "/", " ")
                tx3 = Replace(tx3, " ", " " & tx2)
                CircuitNumber.Offset(0, 1).Value = tx3 & " " & rating & " " & type1
        
        '*** MCB's TP, mcbs triple pole **  WITH EMergency Lights  ** circuits
        '**********************************************************************
            ElseIf CNLenght > 6 And UCase(CircuitNumber.Columns(4).Value) = "MCB" And InStr(islight, "lighting") <> 0 Then ' if total number of characters are bigger then 6, L1/L2/L3/8  to 8L1 8L2 8L3
                mcb3plight = Split(tx1, "/") ' split L1/L2/L3/ in three strings stored in mcb3plight array ('L1' 'L2' 'L3')
                
                tx3 = "E" & Enumber & " " & tx2 & mcb3plight(0) & " " & rating & " " & type1 & " " ' form string E1 1L1 10A 'C'
                    Enumbercount = Enumbercount + 1
                    If Enumbercount > 4 Then
                        Enumber = Enumber + 1
                        Enumbercount = 1
                    End If
                tx3 = tx3 & "E" & Enumber & " " & tx2 & mcb3plight(1) & " " & rating & " " & type1 & " " 'form string E1 1L1 10A 'C' and add E1 1L2 10A 'C'
                    Enumbercount = Enumbercount + 1
                    If Enumbercount > 4 Then
                        Enumber = Enumber + 1
                        Enumbercount = 1
                    End If
                tx3 = tx3 & "E" & Enumber & " " & tx2 & mcb3plight(2) & " " & rating & " " & type1 & " " 'form string E1 1L1 10A 'C' and add E1 1L2 10A 'C' and E1 1L3 10A 'C'
                    Enumbercount = Enumbercount + 1
                    If Enumbercount > 4 Then
                        Enumber = Enumber + 1
                        Enumbercount = 1
                    End If
            CircuitNumber.Offset(0, 1).Value = tx3 ' set value on cell like ¦E2 9L1 10A 'B' E2 9L2 10A 'B' E2 9L3 10A 'B'¦
            
        '*** RCBO's DP, RCBO's double pole
            ElseIf CNLenght <= 6 And UCase(CircuitNumber.Columns(4).Value) = "RCBO" Then
                CircuitNumber.Offset(0, 1).Value = tx3 & " " & tx3 & "N" & " " & rating & " " & type1
                
'        '*** RCBO's DP, RCBO's double pole WITH LIGHTS
'        '**********************************************
'            ElseIf CNLenght <= 6 And UCase(CircuitNumber.Columns(4).Value) = "RCBO" And InStr(islight, "lighting") <> 0 Then
'                CircuitNumber.Offset(0, 1).Value = tx3 & " " & tx3 & "N" & " " & rating & " " & type1
            
        '*** RCBO's TP, RCBO's triple pole
           'ElseIf CNLenght > 6 And UCase(CircuitNumber.Columns(4).Value) = "RCBO" And InStr(islight, "lighting") = 0 Then
            ElseIf CNLenght > 6 And UCase(CircuitNumber.Columns(4).Value) = "RCBO" Then
                tx3 = Replace(tx3, "/", " ")
                tx3 = Replace(tx3, " ", " " & tx2)
                CircuitNumber.Offset(0, 1).Value = tx3 & " " & rating & " " & type1
                
'        '*** RCBO's TP, RCBO's triple pole WITH LIGHTS
'        '*************************************************
'            ElseIf CNLenght > 6 And UCase(CircuitNumber.Columns(4).Value) = "RCBO" And InStr(islight, "lighting") <> 0 Then
'                tx3 = Replace(tx3, "/", " ")
'                tx3 = Replace(tx3, " ", " " & tx2)
'                CircuitNumber.Offset(0, 1).Value = tx3 & " " & rating & " " & type1
                
            End If
        End If

    Next CircuitNumber

'*******create text file with script**************

Dim FSO As Object
Dim TextFile As Object
Dim dist As Integer ' distance betwenn blocks
Dim k1, k2 As Integer  ' variables in 'for next' loops
Dim LArray() As String ' LArray() array to get devices tags splited by space
Dim deviceslist(1 To 200) As String ' deviceslist() arra list of devices
Dim mcb_sp, mcb_sp_E, mcb_sp_E_C, mcb_tp, rcbo_dp, rcbo_tp As String ' dwg block paths

'**** define dwg blocks paths
mcb_sp = """\\server1\AutoCAD\FILES\Jorge\CADblocks\Mcbs_Rcbos_DistBoards\MCB sp.dwg""" & vbCrLf ' mcb single pole
mcb_sp_E = """\\server1\AutoCAD\FILES\Jorge\CADblocks\Mcbs_Rcbos_DistBoards\MCB sp_EmergencyLight.dwg""" & vbCrLf ' mcb single pole with emergency contactor
mcb_sp_E_C = """\\server1\AutoCAD\FILES\Jorge\CADblocks\Mcbs_Rcbos_DistBoards\MCB sp_EmergencyLight_Contactor.dwg""" & vbCrLf ' mcb single pole with emergency contactor and light contactor
mcb_tp = """\\server1\AutoCAD\FILES\Jorge\CADblocks\Mcbs_Rcbos_DistBoards\MCB tp.dwg""" & vbCrLf ' mcb tripole
rcbo_dp = """\\server1\AutoCAD\FILES\Jorge\CADblocks\Mcbs_Rcbos_DistBoards\RCBO dp.dwg""" & vbCrLf ' rcbo double pole
rcbo_tp = """\\server1\AutoCAD\FILES\Jorge\CADblocks\Mcbs_Rcbos_DistBoards\RCBO tp.dwg""" & vbCrLf ' rcbo tripole
rcbo_dp_c = """\\server1\AutoCAD\FILES\Jorge\CADblocks\Mcbs_Rcbos_DistBoards\RCBO dp Contactor.dwg""" & vbCrLf ' rcbo double pole with contactor
rcbo_tp_c = """\\server1\AutoCAD\FILES\Jorge\CADblocks\Mcbs_Rcbos_DistBoards\RCBO tp Contactor.dwg""" & vbCrLf ' rcbo tripole with contactor

'****path to write textfile
TextFilePath = "\\server1\AutoCAD\FILES\Jorge\scripts Autocad\0_AutoDistBoards.scr"


Set FSO = CreateObject("Scripting.FileSystemObject")
dist = 0
k2 = 0

Set TextFile = FSO.CreateTextFile(TextFilePath, True, True)
    
        For Each devicerange In Range("d1:d200") ' make a list of devices to setup best distance between
            deviceslist(devicerange.Row) = devicerange.Value
        Next devicerange
    
    For Each current In Range("b1:b200")
    islight = LCase(current.Columns(5).Value) ' set string value as load name device to check if it contain the word  "lighting"
    
        If IsEmpty(current.Value) = False Then
                LArray = Split(current.Value)
            
                '*** MCB's SP, mcbs single pole no lighting
                If Len(current.Columns(0)) <= 6 And UCase(current.Columns(3).Value) = "MCB" And InStr(islight, "lighting") = 0 Then
                    dist = dist + 130
                    TextFile.Write "-INSERT" & vbCrLf
                    TextFile.Write mcb_sp
                    TextFile.Write "*" & dist & ",0,0" & vbCrLf
                    TextFile.Write "1" & vbCrLf
                    TextFile.Write "1" & vbCrLf
                    TextFile.Write "0" & vbCrLf
                    TextFile.Write LArray(0) & vbCrLf
                    TextFile.Write LArray(1) & vbCrLf
                    TextFile.Write LArray(2) & vbCrLf
                    
                
                '*** MCB's SP, mcbs single pole WITH LIGHTNING contactor
                ElseIf Len(current.Columns(0)) <= 6 And UCase(current.Columns(3).Value) = "MCB" And InStr(islight, "lighting") <> 0 Then
                    dist = dist + 210
                    TextFile.Write "-INSERT" & vbCrLf
                    TextFile.Write mcb_sp_E
                    TextFile.Write "*" & dist & ",0,0" & vbCrLf
                    TextFile.Write "1" & vbCrLf
                    TextFile.Write "1" & vbCrLf
                    TextFile.Write "0" & vbCrLf
                    TextFile.Write LArray(0) & vbCrLf
                    TextFile.Write LArray(1) & vbCrLf
                    TextFile.Write LArray(2) & vbCrLf
                    TextFile.Write LArray(3) & vbCrLf
                    
                
                '*** MCB's TP, mcbs triple pole
                ElseIf Len(current.Columns(0)) > 6 And UCase(current.Columns(3).Value) = "MCB" And InStr(islight, "lighting") = 0 Then
                    dist = dist + 130
                    TextFile.Write "-INSERT" & vbCrLf
                    TextFile.Write mcb_tp
                    TextFile.Write "*" & dist & ",0,0" & vbCrLf
                    TextFile.Write "1" & vbCrLf
                    TextFile.Write "1" & vbCrLf
                    TextFile.Write "0" & vbCrLf
                    TextFile.Write LArray(0) & vbCrLf
                    TextFile.Write LArray(1) & vbCrLf
                    TextFile.Write LArray(2) & vbCrLf
                    TextFile.Write LArray(3) & vbCrLf
                    TextFile.Write LArray(4) & vbCrLf
                    dist = dist + 130
                    
                
                '*** MCB's TP, mcbs triple pole WITH LIGHTNING contactor
                ElseIf Len(current.Columns(0)) > 6 And UCase(current.Columns(3).Value) = "MCB" And InStr(islight, "lighting") <> 0 Then
                    dist = dist + 210
                    TextFile.Write "-INSERT" & vbCrLf
                    TextFile.Write mcb_sp_E
                    TextFile.Write "*" & dist & ",0,0" & vbCrLf
                    TextFile.Write "1" & vbCrLf
                    TextFile.Write "1" & vbCrLf
                    TextFile.Write "0" & vbCrLf
                    TextFile.Write LArray(0) & vbCrLf
                    TextFile.Write LArray(1) & vbCrLf
                    TextFile.Write LArray(2) & vbCrLf
                    TextFile.Write LArray(3) & vbCrLf
                    
                    dist = dist + 180
                    TextFile.Write "-INSERT" & vbCrLf
                    TextFile.Write mcb_sp_E
                    TextFile.Write "*" & dist & ",0,0" & vbCrLf
                    TextFile.Write "1" & vbCrLf
                    TextFile.Write "1" & vbCrLf
                    TextFile.Write "0" & vbCrLf
                    TextFile.Write LArray(4) & vbCrLf
                    TextFile.Write LArray(5) & vbCrLf
                    TextFile.Write LArray(6) & vbCrLf
                    TextFile.Write LArray(7) & vbCrLf
                    
                    dist = dist + 180
                    TextFile.Write "-INSERT" & vbCrLf
                    TextFile.Write mcb_sp_E
                    TextFile.Write "*" & dist & ",0,0" & vbCrLf
                    TextFile.Write "1" & vbCrLf
                    TextFile.Write "1" & vbCrLf
                    TextFile.Write "0" & vbCrLf
                    TextFile.Write LArray(8) & vbCrLf
                    TextFile.Write LArray(9) & vbCrLf
                    TextFile.Write LArray(10) & vbCrLf
                    TextFile.Write LArray(11) & vbCrLf
                    
                    
                                        
                '*** RCBO's DP, RCBO's double pole
                ElseIf Len(current.Columns(0)) <= 6 And UCase(current.Columns(3).Value) = "RCBO" Then
                    dist = dist + 150
                    TextFile.Write "-INSERT" & vbCrLf
                    TextFile.Write rcbo_dp
                    TextFile.Write "*" & dist & ",0,0" & vbCrLf
                    TextFile.Write "1" & vbCrLf
                    TextFile.Write "1" & vbCrLf
                    TextFile.Write "0" & vbCrLf
                    TextFile.Write LArray(0) & vbCrLf
                    TextFile.Write LArray(1) & vbCrLf
                    TextFile.Write LArray(2) & vbCrLf
                    TextFile.Write LArray(3) & vbCrLf
                    dist = dist + 60
                    
                
                '*** RCBO's TP, RCBO's triple pole
                ElseIf Len(current.Columns(0)) > 6 And UCase(current.Columns(3).Value) = "RCBO" Then
                    dist = dist + 130
                    TextFile.Write "-INSERT" & vbCrLf
                    TextFile.Write rcbo_tp
                    TextFile.Write "*" & dist & ",0,0" & vbCrLf
                    TextFile.Write "1" & vbCrLf
                    TextFile.Write "1" & vbCrLf
                    TextFile.Write "0" & vbCrLf
                    TextFile.Write LArray(0) & vbCrLf
                    TextFile.Write LArray(1) & vbCrLf
                    TextFile.Write LArray(2) & vbCrLf
                    TextFile.Write LArray(3) & vbCrLf
                    TextFile.Write LArray(4) & vbCrLf
                    dist = dist + 130
                    
                    
                End If
        End If
    Next current
    'last part of the script organize blocks
    'TextFile.Write
'    Zoom
'*-130,1500,0
'15000,-260,0
'Move
'*3745,704,0
'15000,0,0
'
'*0,0,0
'0,-570,0
'Move
'*7393,50,0
'11233,-650,0
'
'*0,0,0
'0,-570,0
'
'
'
    
TextFile.Close ' close text file
Set FSO = Nothing

Range("A1").Select

'********************** text command script autocad***********
'-INSERT; command insert
'"C:\Users\jorge\Desktop\Temp_desk\lisp_test\MCB sp.dwg" ;name of the block
'*900,0,0 ;position
'1
'1
'0
'1L55 ;circuitnumber
'10A ;current
''C' ;type

'**************** end text command script autocad*************
End Sub
