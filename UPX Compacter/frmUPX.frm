VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmUPX 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "UPX .EXE Compressor"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3120
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   7485
      Begin VB.ComboBox cbIcon 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmUPX.frx":0000
         Left            =   1800
         List            =   "frmUPX.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1755
         Width           =   5505
      End
      Begin VB.CheckBox chTest 
         Caption         =   "&Test integrity"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1800
         TabIndex        =   10
         Top             =   2205
         Value           =   1  'Checked
         Width           =   1635
      End
      Begin VB.TextBox txPath 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   315
         Width           =   5100
      End
      Begin VB.OptionButton opLevel 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1755
         TabIndex        =   5
         Top             =   1305
         Width           =   465
      End
      Begin VB.OptionButton opLevel 
         Caption         =   "&Use best possible"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1755
         TabIndex        =   4
         Top             =   945
         Value           =   -1  'True
         Width           =   1725
      End
      Begin VB.CommandButton btCompress 
         Caption         =   "&Compress"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   315
         TabIndex        =   3
         Top             =   2565
         Width           =   2355
      End
      Begin VB.CommandButton btCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4950
         TabIndex        =   2
         Top             =   2565
         Width           =   2355
      End
      Begin VB.CommandButton btBrowse 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6930
         TabIndex        =   1
         Top             =   315
         Width           =   375
      End
      Begin MSComDlg.CommonDialog CDL 
         Left            =   5490
         Top             =   2250
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "Locate the program's executable"
         Filter          =   "Executable files|*.dll;*.exe|All files|*.*"
      End
      Begin MSComctlLib.Slider slLevel 
         Height          =   285
         Left            =   2250
         TabIndex        =   6
         Top             =   1305
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   503
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   9
         SelStart        =   8
         Value           =   8
      End
      Begin VB.Label lbIcon 
         Caption         =   "Compress icons:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         TabIndex        =   11
         Top             =   1800
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Path to executable:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   9
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label2 
         Caption         =   "Compression level:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   315
         TabIndex        =   8
         Top             =   945
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmUPX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim sExeName As String 'Full path to the project's executable

'--- Browse for the executable ---
Private Sub btBrowse_Click()
    On Error GoTo CancelBrowse
    With CDL
        .InitDir = GetDirName(sExeName)
        .ShowOpen
        sExeName = .Filename
        txPath.Text = sExeName
    End With
CancelBrowse:
End Sub

'--- remembers the executable's path passed from the VB IDE interface ---
Sub SetProject(sProjectExe As String)
    sExeName = sProjectExe
    txPath.Text = sExeName
End Sub

'--- Closes the form ---
Private Sub btCancel_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub btCompress_Click()
    CompressFile
End Sub

Private Sub Form_Load()
    cbIcon.ListIndex = 2
End Sub

'--- Sets the caption of the radio button when slider's value changes ---
Private Sub slLevel_Change()
    opLevel(0).Caption = slLevel.Value
End Sub

'--- Prepares the parameter string for the UPX.EXE and runs it, then reports to the user ---
Sub CompressFile()
    Dim iLevel As Integer           'compression level
    Dim sParam As String            'holds the upx.exe command line string
    Dim sAppPath As String          'path in which the VB Add-in DLL resides
    Dim sExePath As String          'full path to the executable to be compressed
    Dim sExeShortName As String     'the file name of the executable to be compressed (path stripped)
    
    Dim ret As Long                 'return value from upx.exe
    Dim FSize(1) As Long            'file sizes of the executable before and after compression
    Dim bTest As Boolean            'If true, perform integrity test
    Dim sResult As String           'holds the result from integrity test
    
    On Error GoTo CompressErr
    Screen.MousePointer = 11
    'prepare paths
    sAppPath = ToPath(App.Path)
    sExePath = ToPath(GetDirName(sExeName))
    sExeShortName = GetFileName(sExeName)
    'prepare flags
    bTest = chTest.Value = 1
    
    'check if executable exists
    If Dir(sExeName) = "" Then
        MsgBox "The specified executable " + sExeName + " does not exist.", vbCritical
    Else
        'prepare parameters string
        If opLevel(0) Then 'set compression level
            sParam = "-" + CStr(slLevel.Value)
        Else
            sParam = "--best"
        End If
        'set icon compression option
        sParam = sParam + " --compress-icons#" + CStr(cbIcon.ListIndex)
        'set allocated memory and compression method (there is --nrv2d too)
        sParam = sParam + " --crp-ms=999999 --nrv2b"
        'append the file name
        sParam = sParam + " " + Chr(34) + sExeShortName + Chr(34)
        'measure the file size before compression
        FSize(0) = FileLen(sExeName)
        'copy upx.exe to the executable's path if necessary
        If Dir(sExePath + "upx.exe") = "" Then
            FileCopy sAppPath + "upx.exe", sExePath + "upx.exe"
        End If
        'set current directory to the executable's path
        ChDrive Left(sExePath, 1)
        ChDir sExePath
        'execute UPX.exe and wait for result to be returned
        ret = ShellAndWait("upx.exe " + sParam)
        'check result
        If ret = 0 Then 'success
            FSize(1) = FileLen(sExeName)
            If bTest Then
                ret = ShellAndWait("upx.exe -t " + Chr(34) + sExeShortName + Chr(34))
                If ret <> 0 Then
                    sResult = Choose(ret, "error.", "warning.")
                    MsgBox "Integrity test returned " + sResult, vbExclamation, "Integrity test"
                End If
            End If
            If MsgBox("The file has been downsized from " + Format(FSize(0) / 1024, "#,##0.0") + "K to " + Format(FSize(1) / 1024, "#,##0.0") + "K (" + Format((FSize(1) - FSize(0)) / FSize(0), "0.0%") + ")." + vbCrLf + "Do you wish to run the executable now?", vbYesNo + vbInformation, "Success") = vbYes Then
                ShellExecute Me.hwnd, "open", sExeName, "", "", 1
            End If
        ElseIf ret = 2 Then 'error or already compacted (most probably)
            MsgBox "This file is already compressed.", vbExclamation
        Else 'some warning
            MsgBox "Command failed (Code " + CStr(ret) + ").", vbCritical, "Failure."
        End If
    End If
ExitCompress:
    On Error Resume Next
    'clean-up
    Kill sExePath + "upx.exe"
    'exit procedure
    Screen.MousePointer = 0
    Me.Hide
    Unload Me
    Exit Sub
CompressErr: 'error handler
    AddInErr Err, "Compressing file..."
    Resume ExitCompress
End Sub

'--- Executes a file synchronously; may cause warnings in some antivirus programs! ---
Function ShellAndWait(Filename As String) As Long
    Dim objScript As Object
    Dim ShellApp As Long
    
    On Error GoTo ERR_OpenForEdit
    'use windows scripting host object to run the program
    Set objScript = CreateObject("WScript.Shell")
    ShellApp = objScript.Run(Filename, 1, True)
    ShellAndWait = ShellApp 'return operation result
        
EXIT_OpenForEdit:
    Set objScript = Nothing
    Exit Function
ERR_OpenForEdit:
    GoTo EXIT_OpenForEdit
End Function

