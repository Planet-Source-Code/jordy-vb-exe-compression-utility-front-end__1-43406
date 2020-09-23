Attribute VB_Name = "modMain"
Option Explicit

'*****************************************************************
'*            VB Add-In written by Chavdar Jordanov              *
'*                 (chavdar_jordanov@yahoo.com)                  *
'*         Front-end to the UPX Compression Utility              *
'*            by Markus Oberhumer & Laszlo Molnar                *
'*                 http://upx.sourceforge.net                    *
'*****************************************************************

'--- Error processing code ---
Public Sub AddInErr(obErr As ErrObject, Optional sCaption As String)
    'write to error log
    Open App.Path + "\errors.log" For Append As #1
    Print #1, CStr(Now); " // "; sCaption
    Print #1, "    Error: " + obErr.Description
    Print #1, "====================================="
    Close #1
    'comment the next line if you don't wish any error messages displayed
    MsgBox obErr.Description, vbCritical, "Add-in Error"
End Sub

'--- Appends "\" to the file path ---
Public Function ToPath(ByVal sPath As String) As String
    If sPath <> "" Then If Right(sPath, 1) <> "\" Then sPath = sPath + "\"
    ToPath = sPath
End Function

'--- Extracts the file name from a full path ---
Function GetFileName(sPath As String) As String
    Dim X
    X = InStrRev(sPath, "/")
    If X = 0 Then
        X = InStrRev(sPath, "\")
        If X = 0 Then GetFileName = sPath: Exit Function
    End If
    GetFileName = Mid(sPath, X + 1)
End Function

'--- Extracts directory name from a path ---
Function GetDirName(sPath As String) As String
    Dim X
        X = InStrRev(sPath, "/")
    If X = 0 Then
        X = InStrRev(sPath, "\")
        If X = 0 Then GetDirName = sPath: Exit Function
    End If
    GetDirName = Left(sPath, X - 1)
End Function
