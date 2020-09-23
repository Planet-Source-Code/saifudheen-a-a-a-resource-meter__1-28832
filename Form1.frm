VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Free Resources"
   ClientHeight    =   915
   ClientLeft      =   2760
   ClientTop       =   4155
   ClientWidth     =   4470
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   61
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   298
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3000
      Top             =   30
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Hide"
      Height          =   345
      Left            =   3510
      TabIndex        =   1
      Top             =   210
      Width           =   885
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawWidth       =   2
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2490
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   285
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3660
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   8421504
      _Version        =   393216
   End
   Begin VB.Menu mnuShow 
      Caption         =   "Show"
      Visible         =   0   'False
      Begin VB.Menu mnuRes 
         Caption         =   "Free &System Resource"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuRes 
         Caption         =   "Free &GDI Resource"
         Index           =   2
      End
      Begin VB.Menu mnuRes 
         Caption         =   "Free &User Resource"
         Index           =   3
      End
      Begin VB.Menu mnuRes 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuRes 
         Caption         =   "Show..."
         Index           =   5
      End
      Begin VB.Menu mnuRes 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuRes 
         Caption         =   "&About..."
         Index           =   7
      End
      Begin VB.Menu mnuRes 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuRes 
         Caption         =   "E&xit"
         Index           =   9
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'// Free Resources Indicator/Meter
'// Programmed by A. A. Saifudheen. (keraleeyan@msn.com)
'// It will be a great reward to me if you send a message,
'// If you learned something from this code.
'// Please vote for me if you like this code.

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

'constants required by Shell_NotifyIcon API call:
    Private Const NIM_ADD = &H0
    Private Const NIM_MODIFY = &H1
    Private Const NIM_DELETE = &H2
    Private Const NIF_MESSAGE = &H1
    Private Const NIF_ICON = &H2
    Private Const NIF_TIP = &H4
    Private Const WM_MOUSEMOVE = &H200
    Private Const WM_LBUTTONDOWN = &H201     'Button down  'NOT USED
    Private Const WM_LBUTTONUP = &H202       'Button up    'N U
    Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
    Private Const WM_RBUTTONDOWN = &H204     'Button down  'N U
    Private Const WM_RBUTTONUP = &H205       'Button up
    Private Const WM_RBUTTONDBLCLK = &H206   'Double-click 'N U
'constants required by GetFreeResources API
    Const GFSR_SYSTEMRESOURCES = 0
    Const GFSR_GDIRESOURCES = 1
    Const GFSR_USERRESOURCES = 2
    
    Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
    Private Declare Function GetFreeResources Lib "RSRC32" Alias "_MyGetFreeSystemResources32@4" (ByVal lWhat As Long) As Long
    
    Private nid As NOTIFYICONDATA
    Dim userBox As RECT
    Dim gdiBox As RECT
    Dim systemBox As RECT
    Dim systemFree As Integer    'Storing Value of free resource
    Dim userFree As Integer
    Dim GDIFree As Integer

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    ret = Shell_NotifyIcon(NIM_ADD, nid)
    Me.Hide
End Sub

Private Sub Form_Load()
    'Initialise
    Me.ScaleMode = vbPixels
    With Picture1
        .AutoRedraw = True
        .ScaleMode = vbPixels
        '.Width = 15
       ' .Height = 15
        .BackColor = vbWhite
    End With
    ImageList1.MaskColor = vbWhite   'set transparent color for ExtractICon
    
    Dim ic As Picture
    ImageList1.ListImages.Add 1, , Picture1.Image 'initially add a picture to imagelist
    Set ic = CreateIcon(Picture1.Image)

    With nid
     .cbSize = Len(nid)
     .hwnd = Me.hwnd
     .uId = vbNull
     .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
     .uCallBackMessage = WM_MOUSEMOVE
     .hIcon = ic
    End With
    
    ret = Shell_NotifyIcon(NIM_ADD, nid) 'Create New Icon on Task bar
    Me.Hide
    ShowResource
    With systemBox  'These 3 Rects  are for Indicators in Form
        .left = 50
        .top = 5
        .right = 200
        .bottom = 15
    End With
    With gdiBox
        .left = 50
        .top = 25
        .right = 200
        .bottom = 35
    End With
    With userBox
        .left = 50
        .top = 45
        .right = 200
        .bottom = 55
    End With
End Sub

Private Function CreateIcon(img As Picture) As Picture  'Bitmap to Icon
    'This code converts a bitmap picture to an Icon using an ImageList controll
    ' as ShellNotifyIcon will not work without an Icon
    ImageList1.ListImages.Remove 1
    ImageList1.ListImages.Add 1, , img
    Set CreateIcon = ImageList1.ListImages(1).ExtractIcon
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Result As Long
    Dim msg As Long
    msg = X
    Select Case msg
        Case WM_LBUTTONDBLCLK
            KillIcon 'Delete icon from taskbar
            Me.Show
        Case WM_RBUTTONUP
            Me.PopupMenu mnuShow
    End Select

End Sub

Private Sub mnuExit_Click()
    Unload Form1
End Sub


Private Sub Form_Paint()
    Dim user As RECT
    Dim gdi As RECT
    Dim system As RECT  'these 3 rect structurs are used for indicators
    user = userBox
    gdi = gdiBox
    system = systemBox
    
    system.right = system.left + (systemBox.right - systemBox.left) * systemFree / 100
    Me.Line (system.left, system.top)-(system.right, system.bottom), QBColor(8), BF

    gdi.right = gdi.left + (gdiBox.right - gdiBox.left) * GDIFree / 100
    Me.Line (gdi.left, gdi.top)-(gdi.right, gdi.bottom), QBColor(5), BF

    user.right = user.left + (userBox.right - userBox.left) * userFree / 100
    Me.Line (user.left, user.top)-(user.right, user.bottom), vbBlue, BF

    Me.Line (userBox.left, userBox.top)-(userBox.right, userBox.bottom), , B
    Me.Line (gdiBox.left, gdiBox.top)-(gdiBox.right, gdiBox.bottom), , B
    Me.Line (systemBox.left, systemBox.top)-(systemBox.right, systemBox.bottom), , B
    
    Me.CurrentY = systemBox.top
    Me.CurrentX = 3
    Me.Print "System :"
    Me.CurrentY = systemBox.top
    Me.CurrentX = systemBox.right + 3
    Me.Print systemFree & " %"
    
    Me.CurrentY = gdiBox.top
    Me.CurrentX = 3
    Me.Print "GDI :"
    Me.CurrentY = gdiBox.top
    Me.CurrentX = gdiBox.right + 3
    Me.Print GDIFree & " %"
    
    Me.CurrentY = userBox.top
    Me.CurrentX = 3
    Me.Print "User :"
    Me.CurrentY = userBox.top
    Me.CurrentX = userBox.right + 3
    Me.Print userFree & " %"
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Shell_NotifyIcon NIM_ADD, nid ' Every time form is minimized;
                                        'Icon is created on TaskBar
    Else
        KillIcon                       'and when form is visible Icon is deleted
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    KillIcon
End Sub

Private Sub KillIcon()
    Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub mnuRes_Click(Index As Integer)
    If Index = 5 Then  'Menu Show Form
        KillIcon
        Me.Show
        Exit Sub
    End If
    If Index = 7 Then 'about
        MsgBox "Resource Indicator." & vbCrLf & "Programmed by A. A. Saifudheen." & vbCrLf & "keraleeyan@msn.com.", vbOKOnly, "About"
        Exit Sub
    End If
    If Index = 9 Then  'Menu Exit
        Unload Me
        Exit Sub
    End If
    
    mnuRes(Index).Checked = True
    For i = 1 To 3
       If Index <> i Then mnuRes(i).Checked = False
    Next i
    ShowResource
End Sub

Private Sub Timer1_Timer()
    ShowResource
End Sub

Private Sub ShowResource()
    Dim Selected As String
    Dim Data As Integer
    Static count As Integer
    count = count + 1  'used for blinking the icon on low resource
    If count = 10 Then count = 0
    
    systemFree = GetFreeResources(GFSR_SYSTEMRESOURCES)
    GDIFree = GetFreeResources(GFSR_GDIRESOURCES)
    userFree = GetFreeResources(GFSR_USERRESOURCES)
    
    Select Case True
    Case mnuRes(1).Checked  'System resource
        Selected = "Free System Resource %"
        Picture1.ForeColor = QBColor(0)
        Data = systemFree
    Case mnuRes(2).Checked  'GDI
        Selected = "Free GDI Resource %"
        Picture1.ForeColor = QBColor(5)
        Data = GDIFree
    Case mnuRes(3).Checked  'User
        Selected = "Free User Resource %"
        Picture1.ForeColor = QBColor(9)
        Data = userFree
    End Select
    
    If Val(Data) < 10 Then  'Very low free resource (red)
        Picture1.ForeColor = QBColor(4)
    End If
    

    'These two lines actually makes a bitmap image of the display
    Picture1.Cls
    Picture1.Print Data
    
    If systemFree < 10 Or GDIFree < 10 Or userFree < 10 Then
        If count Mod 2 = 0 Then Picture1.Cls  'If any resource is on low
                                                'then icon will blink
    End If                                     'This can be tested by loading
                                                'more programms.
    Dim dIcon As Picture
    Set dIcon = CreateIcon(Picture1.Image)
    With nid
     .hIcon = dIcon
     .szTip = Selected & vbNullChar
    End With
    ret = Shell_NotifyIcon(NIM_MODIFY, nid)
    Me.Cls
    Form_Paint
End Sub
