Option Explicit
Option Base 1

Private Type UINT64
    LowPart As Long
    HighPart As Long
End Type
Private Const BSHIFT_32 = 4294967296# ' 2 ^ 32

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As UINT64) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As UINT64) As Long
  
'Variables for timers
Dim uStart As UINT64
Dim uEnd As UINT64
Dim uFreq As UINT64
Dim dblElapsed As Double

Private Function U64Dbl(U64 As UINT64) As Double
    Dim lDbl As Double, hDbl As Double
    lDbl = U64.LowPart
    hDbl = U64.HighPart
    If lDbl < 0 Then lDbl = lDbl + BSHIFT_32
    If hDbl < 0 Then hDbl = hDbl + BSHIFT_32
    U64Dbl = lDbl + BSHIFT_32 * hDbl
End Function

Sub Pomiar_Start()
    QueryPerformanceFrequency uFreq
    QueryPerformanceCounter uStart
End Sub

Sub Pomiar_Koniec()
    QueryPerformanceCounter uEnd
    Debug.Print Format(Now, "hh") & ":" & Format(Now, "Nn") & ": " & Format((U64Dbl(uEnd) - U64Dbl(uStart)) / U64Dbl(uFreq), "0.000000"); " seconds elapsed."
End Sub

'================================================================================
' RestoreCoruptedCSV
'
' Author: Marcin H
' Version: early beta v 0.1
' Date: 18/11/2014
' Last Update: 21/11/2014
'
' Restore structure of corupted CSV file based on count of records from first row. It was designed to
' restore order in corupted by unnecesary CrLf (Enter sign) in records.
' Works for almost 100% cases except one...
'
' Need to be done:
' 1. ANSI vs Unicode checker
' 2. Unnecessary CrLf in last column of data
' 3. Type of delimiter checker f.e. Tab or ","
'================================================================================

Sub RestoreCoruptedCSV()

Dim strPath As String
Dim strFile As String
Dim strFileNew As String
Dim strDir As String

Dim strLine As String
Dim strLineTemp As String
Dim sTemp As String

Dim objFSO As Scripting.FileSystemObject
Dim objTextFile As TextStream
Dim objToWriteTxt As TextStream

Dim lMemberCountLine As Long, lMemberCount As Long, i As Long

Set objFSO = CreateObject("Scripting.FileSystemObject")

' 1. position from "Need to be done"
strPath = Application.GetOpenFilename(FileFilter:="Pliki tekstowe (*.csv;*.txt;),*.csv;*.txt;", _
                                      Title:="Wybierz plik tekstowy")

If objFSO.FileExists(strPath) = False Then
    MsgBox "No file was selected."
    End
End If

Pomiar_Start

Set objTextFile = objFSO.OpenTextFile(strPath, ForReading, False, TristateTrue)

strDir = objFSO.GetParentFolderName(strPath) & "\"
strFileNew = "new(1)_" & objFSO.GetFileName(strPath)

If objFSO.FileExists(strDir & strFileNew) Then
    sTemp = strFileNew
    Do Until objFSO.FileExists(strDir & sTemp) = False
        i = CInt(Mid(sTemp, InStr(sTemp, "(") + 1, InStr(sTemp, ")") - InStr(sTemp, "(") - 1)) + 1
        sTemp = Left(sTemp, InStr(sTemp, "(")) & i & Right(sTemp, Len(sTemp) - InStr(sTemp, ")") + 1)
    Loop
    Set objToWriteTxt = objFSO.CreateTextFile(strDir & sTemp, False, True)
Else
    Set objToWriteTxt = objFSO.CreateTextFile(strDir & strFileNew, False, True)
End If

With objTextFile
    'First line defines number of members in delimited file.
    strLine = .ReadLine
    lMemberCount = UBound(Split(strLine, ","))
    objToWriteTxt.WriteLine (strLine)
    
    Do Until objTextFile.AtEndOfStream
        strLine = .ReadLine
        lMemberCountLine = UBound(Split(strLine, ","))
        If lMemberCountLine = lMemberCount Then
            objToWriteTxt.WriteLine (strLine)
            strLine = vbNullString
        Else
            Do While lMemberCountLine < lMemberCount
                strLineTemp = .ReadLine
                strLine = strLine & strLineTemp
                lMemberCountLine = UBound(Split(strLine, ","))
                If lMemberCountLine >= lMemberCount Then
                    objToWriteTxt.WriteLine (strLine)
                    strLine = vbNullString
                End If
            Loop
        End If
        i = i + 1
        If i Mod 100 = 0 Then
            DoEvents
        End If
    Loop
    .Close
End With
objToWriteTxt.Close

Pomiar_Koniec

End Sub
