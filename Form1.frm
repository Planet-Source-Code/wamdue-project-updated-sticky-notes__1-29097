VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmNote 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Note"
   ClientHeight    =   3030
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4140
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":000C
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   276
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   705
      ItemData        =   "Form1.frx":2BB36
      Left            =   360
      List            =   "Form1.frx":2BB43
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Timer tmrEvent 
      Interval        =   15000
      Left            =   0
      Top             =   2640
   End
   Begin MSComCtl2.DTPicker myTime 
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   65535
      CalendarForeColor=   255
      CalendarTitleBackColor=   65535
      CalendarTitleForeColor=   255
      CalendarTrailingForeColor=   65535
      Format          =   22740994
      CurrentDate     =   37216
   End
   Begin MSComCtl2.DTPicker myDate 
      Height          =   255
      Left            =   840
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   65535
      CalendarForeColor=   0
      CalendarTitleBackColor=   65535
      CalendarTitleForeColor=   0
      CalendarTrailingForeColor=   65535
      Format          =   22740993
      CurrentDate     =   37216
   End
   Begin VB.TextBox txtTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   840
      MaxLength       =   12
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Sticky Notes"
      Top             =   240
      Width           =   2055
   End
   Begin RichTextLib.RichTextBox txtNote 
      Height          =   2055
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3625
      _Version        =   393217
      BackColor       =   65535
      BorderStyle     =   0
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Form1.frx":2BB85
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "&Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblMinimise 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   135
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   200
      X2              =   208
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Label lblClose 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   135
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   216
      X2              =   224
      Y1              =   24
      Y2              =   16
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   216
      X2              =   224
      Y1              =   16
      Y2              =   24
   End
End
Attribute VB_Name = "frmNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TimeChanged As Boolean
Dim bCtrl As Boolean



Private Sub Form_Load()

myDate.Value = Date
tmrEvent.Enabled = False

FirstOpen = True
lstMenu.Selected(2) = True
FirstOpen = False

If Me.Picture <> 0 Then
  Call SetAutoRgn(Me)
End If

End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
lstMenu.Visible = False
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub Label1_Click()
    lstMenu.Visible = True
End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub lblMinimise_Click()
    
    If TimeChanged = True Then
        TimeChanged = False
        tmrEvent.Enabled = True
    End If
    
    If lstMenu.Selected(2) = False Then
        Me.WindowState = vbMinimized
    Else
        tmrEvent.Enabled = True
        Me.Visible = False
    End If
    
    lstMenu.Visible = False
    
End Sub

Private Sub lstMenu_Click()
    
    If FirstOpen = True Then
        FirstOpen = False
        Exit Sub
    End If
    
    Select Case lstMenu.ListIndex
        Case 0
            If lstMenu.Selected(0) = True Then
                Me.myTime.Visible = True
                Me.myDate.Visible = True
                'set textbox height
                Me.txtNote.Height = 105
            Else
                Me.myTime.Visible = False
                Me.myDate.Visible = False
                'turn off timer if it's on
                Me.tmrEvent.Enabled = False
                'set textbox height
                Me.txtNote.Height = 129
            End If
        Case 1
            If lstMenu.Selected(1) = True Then
                rtn = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
            Else
                rtn = SetWindowPos(Me.hWnd, -2, 0, 0, 0, 0, 3)
            End If

    End Select
    
    lstMenu.Visible = False
    
End Sub

Private Sub lstMenu_LostFocus()
    lstMenu.Visible = False
End Sub

Private Sub myTime_Change()
    TimeChanged = True
End Sub

Private Sub tmrEvent_Timer()
    If myDate = Date And Format(myTime, "hh:mm") = Format(Time, "hh:mm") Then
        Me.Show
        rtn = SetWindowPos(Me.hWnd, -1, 0, 0, 0, 0, 3)
        tmrEvent.Enabled = False
    End If
End Sub



Private Sub txtNote_Click()
    lstMenu.Visible = False
End Sub

Private Sub txtNote_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 17 Then
    bCtrl = True
 Else
     If KeyCode <> 77 Then bCtrl = False
 End If
 
End Sub

Private Sub txtNote_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 77 And bCtrl = True Then
        adjustx = Screen.TwipsPerPixelX
        adjusty = Screen.TwipsPerPixelY
        SetCursorPos (Me.Left / adjustx) + Label1.Left, (Me.Top / adjusty) + (Label1.Top * 4)
        Label1_Click
        bCtrl = False
    Else
        If KeyCode = 109 Then
            lblMinimise_Click
        End If
    End If
End Sub

Private Sub txtTitle_Change()
    Me.Caption = txtTitle.Text
End Sub


Private Sub txtTitle_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        txtNote.SetFocus
    End If
End Sub
