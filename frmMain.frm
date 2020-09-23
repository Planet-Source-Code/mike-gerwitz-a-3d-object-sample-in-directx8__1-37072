VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "3D Object Example In DirectX8"
   ClientHeight    =   5820
   ClientLeft      =   1965
   ClientTop       =   1935
   ClientWidth     =   6465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6465
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupTeapot 
         Caption         =   "Create Teapot"
      End
      Begin VB.Menu mnuPopupBox 
         Caption         =   "Create Box"
      End
      Begin VB.Menu mnuPopupSphere 
         Caption         =   "Create Sphere"
      End
      Begin VB.Menu mnuPopupTorus 
         Caption         =   "Create Torus"
      End
      Begin VB.Menu mnuPopupDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupMaxRestore 
         Caption         =   "Maximize/Restore"
      End
      Begin VB.Menu mnuPopupFrameRate 
         Caption         =   "Display/Hide Frame Rate"
      End
      Begin VB.Menu mnuPopupReset 
         Caption         =   "Reset"
      End
      Begin VB.Menu mnuPopupDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupLight 
         Caption         =   "Change Light Color"
      End
      Begin VB.Menu mnuPopupBack 
         Caption         =   "Change Background Color"
      End
      Begin VB.Menu mnuPopupDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuPopupAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuPopupExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Then ChangeView -0.5, 0, 0
    If KeyCode = vbKeyRight Then ChangeView 0.5, 0, 0
    
    If KeyCode = vbKeyUp Then ChangeView 0, -0.5, 0
    If KeyCode = vbKeyDown Then ChangeView 0, 0.5, 0
    
    If KeyCode = vbKeyAdd Then ChangeView 0, 0, 0.5, True
    If KeyCode = vbKeySubtract Then ChangeView 0, 0, -0.5, True
    
    If KeyCode = vbKeyNumpad4 Then Rotate pi / 100
    If KeyCode = vbKeyNumpad6 Then Rotate -pi / 100
End Sub

Private Sub Form_Load()
    Me.Show
    DoEvents
    InitGame
    PlayGame
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then PopupMenu mnuPopup
End Sub

Private Sub mnuPopupAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuPopupBack_Click()
    On Error Resume Next
    
    Dim Num As Long
    Num = InputBox("Please enter the hexidecimal number below", "Change Background Color", BackColor)
    ChangeBackColor Num
End Sub

Private Sub mnuPopupBox_Click()
    frmBox.Show
End Sub

Private Sub mnuPopupExit_Click()
    EndGame
End Sub

Private Sub mnuPopupFrameRate_Click()
    ShowFrameRate = Not ShowFrameRate
End Sub

Private Sub mnuPopupHelp_Click()
    MsgBox "1) To turn and move the object up and down, use the arrow keys" & vbCrLf & _
                 "2) To move away or step forward, press the + and - keys" & vbCrLf & _
                 "3) To turn left, press 4 on the number pad" & vbCrLf & _
                 "4) To turn right, press 6 on the number pad" & vbCrLf & _
                 "3) For more help, refer to the windows displayed when creating the D3D objects", vbInformation, "Help"
End Sub

Private Sub mnuPopupLight_Click()
    frmChangeLight.Show
End Sub

Private Sub mnuPopupMaxRestore_Click()
    If Me.WindowState = 2 Then
        Me.WindowState = 0
    Else
        Me.WindowState = 2
    End If
End Sub

Private Sub mnuPopupReset_Click()
    CreateTeapot
    ViewPoint = vec(0, 0, 0)
    CameraPoint = vec(5, 5, 5)
End Sub

Private Sub mnuPopupSphere_Click()
    frmSphere.Visible = True
End Sub

Private Sub mnuPopupTeapot_Click()
    CreateTeapot
End Sub

Private Sub mnuPopupTorus_Click()
    frmTorus.Show
End Sub
