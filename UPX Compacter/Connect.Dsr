VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} VBConnect 
   ClientHeight    =   9510
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   11400
   _ExtentX        =   20108
   _ExtentY        =   16775
   _Version        =   393216
   Description     =   "Compresses the program's executable using the UPX compressing utility"
   DisplayName     =   "UPX Compressor"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   SatName         =   "irmclient"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "VBConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public VbInstance As VBIDE.VBE      'Reference to the VB IDE currently open
          
Private btOpen As Office.CommandBarButton      'Reference to the button on the application's toolbar
Attribute btOpen.VB_VarHelpID = -1
Public WithEvents MenuHandler As CommandBarEvents 'Event handler for the button
Attribute MenuHandler.VB_VarHelpID = -1

Implements IDTExtensibility2    'Some additional events

Const cst_ToolbarName = "Standard"  'the name of the toolbar to which the add-in button is added
Const cst_Caption = "Compact"       'The caption of the add-in button
Const cst_ToolTip = "Compact executable using the UPX utility" 'tooltip text for the add-in button

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
'-----------------------------------------------------
'we need the following stubs, even though they are empty!
'-----------------------------------------------------
Private Sub IDTExtensibility2_OnAddInsUpdate(custom() As Variant)
    'some comments to prevent compiler from removing the procedure stub
End Sub

Private Sub IDTExtensibility2_OnBeginShutdown(custom() As Variant)
    'some comments to prevent compiler from removing the procedure stub
End Sub

'--- Executes when the VB IDE is open ---
Private Sub IDTExtensibility2_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    'remember the application instance
    Set VbInstance = Application
    'Create a button on the application's toolbar
    Set btOpen = AddToAddInCommandBar(cst_Caption, cst_ToolTip)
    'Set the event handler for the button
    Set Me.MenuHandler = VbInstance.Events.CommandBarEvents(btOpen)
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub IDTExtensibility2_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    'delete the button from the toolbar
    btOpen.Delete
    Set btOpen = Nothing
    'destroy the VB object
    Set VbInstance = Nothing
End Sub

Private Sub IDTExtensibility2_OnStartupComplete(custom() As Variant)
    'some comments to prevent compiler from removing the procedure stub
End Sub


'------------------------------------------------------
'Code for the button click action
'------------------------------------------------------
Sub DoButtonAction()
    Dim sExePath As String
    Dim CurrentProject As VBIDE.VBProject
    On Error GoTo ErrHandler
    'get reference to the currently open project
    Set CurrentProject = VbInstance.ActiveVBProject
    If CurrentProject Is Nothing Then 'the project seems not to be loaded
        MsgBox "Please, open a VB project.", vbExclamation
    Else
        With CurrentProject
            If .IsDirty Then .MakeCompiledFile 're-compile if code has changed
            sExePath = .BuildFileName 'get the path to the executable
        End With
        frmUPX.SetProject sExePath 'set the path to the executable in the main form
        frmUPX.Show 1
        Set CurrentProject = Nothing
    End If
ExitSub:
    Exit Sub
ErrHandler:
    AddInErr Err
    Resume ExitSub
End Sub

'--- Adds the add-in button to the VB's standard toolbar ---
Function AddToAddInCommandBar(sCaption As String, sToolTip As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarButton  'command bar object
    Dim cbMenu As Object
    Dim cPic As StdPicture
    On Error GoTo AddToAddInCommandBarErr
    
    'see if we can find the Add-Ins menu
    Set cbMenu = VbInstance.CommandBars(cst_ToolbarName)
    If cbMenu Is Nothing Then
        'not available so we fail
        Exit Function
    End If
    
    'add it to the command bar
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    'set the caption and other button attributes
    cbMenuCommandBar.Caption = sCaption
    cbMenuCommandBar.ToolTipText = sToolTip
    cbMenuCommandBar.Visible = True
    cbMenuCommandBar.Style = msoButtonIconAndCaption
    'paint the button picture (loaded from disk)
    Set cPic = LoadPicture(ToPath(App.Path) + "face.bmp")
    Clipboard.Clear
    Clipboard.SetData cPic, 2
    cbMenuCommandBar.PasteFace
    'destroy objects
    Set AddToAddInCommandBar = cbMenuCommandBar
    Set cPic = Nothing
    Exit Function
    
AddToAddInCommandBarErr:
    
End Function

'--- Fires when the add-in button is clicked ---
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    DoButtonAction
End Sub
