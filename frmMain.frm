VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   1200
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   1920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MousePointer    =   5  'Size
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   1200
   ScaleWidth      =   1920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   30
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1800
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   165
      Left            =   40
      MousePointer    =   1  'Arrow
      Picture         =   "frmMain.frx":6BCA
      Top             =   40
      Width           =   150
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      X1              =   0
      X2              =   1440
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The following code is only used to allow form draging
'from any part of it

Option Explicit


Private hRgn As Long

'Constants declaration needed for the CommonDialog
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_PATHMUSTEXIST = &H800
Private Const CC_FULLOPEN = &H2
Private Const CC_SOLIDCOLOR = &H80
Private Const CC_RGBINIT = &H1
Private Const CC_ANYCOLOR = &H100

Private Sub Form_Deactivate()
SetWindowPos hwnd, conHwndTopmost, 100, 100, 400, 141, conSwpNoActivate Or conSwpShowWindow
End Sub

Private Sub Form_LostFocus()
SetWindowPos hwnd, conHwndTopmost, 100, 100, 400, 141, conSwpNoActivate Or conSwpShowWindow
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub


Private Sub Form_Load()
    
'transparent color is white..
    CommonDialog1.Color = vbWhite
    SetRegion

    
End Sub

Private Sub Form_Paint()
SetWindowPos hwnd, conHwndTopmost, 100, 100, 140, 80, conSwpNoActivate Or conSwpShowWindow
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Free the used memory by the Region and unload the
'Form
    If hRgn Then DeleteObject hRgn
    
End Sub





Private Sub SetRegion()
'Free the memory set
    If hRgn Then DeleteObject hRgn
'Scan the Bitmap and remove all transparent pixels from
'it, creating a new region
    hRgn = GetBitmapRegion(frmMain.Picture, CommonDialog1.Color)
'Set the Forms new Region
    SetWindowRgn frmMain.hwnd, hRgn, True
End Sub

Private Sub Image1_Click()
Unload Me
'Free the used memory by the Region and unload the
'Form
    If hRgn Then DeleteObject hRgn
End Sub
