VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main Form"
   ClientHeight    =   195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   2325
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   195
   ScaleWidth      =   2325
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtSave 
      Height          =   2775
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   5535
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuNew 
         Caption         =   "New Sticky Note"
      End
      Begin VB.Menu mnuShowAll 
         Caption         =   "Show All Sticky Notes"
      End
      Begin VB.Menu mnuNotesAv 
         Caption         =   "Notes Available"
         Begin VB.Menu mnuSN 
            Caption         =   "Sticky Notes 1"
            Index           =   0
         End
      End
      Begin VB.Menu mnubar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    
    Dim strMyToolTip As String
    
    strMyToolTip = "Sticky Notes"
    Call AddSystray(Me, strMyToolTip)
    
    'check if we have old data from previous sticky notes to show...
    If Dir(App.Path & "\SND.osn", vbNormal) <> "" Then
        'have file
        Dim iFileNumber As Integer
        Dim MyData As Variant
        Dim ReadFile As String
        iFileNumber = FreeFile
        Open App.Path & "\SND.osn" For Input As #iFileNumber
        ReadFile = Input(LOF(iFileNumber), iFileNumber)
        MyData = Split(ReadFile, "#~!@")
        Dim TotalForms As Integer
        For i = 0 To UBound(MyData) - 1
            TotalForms = Forms.Count + 1
            Dform.CreateForm "Form" + CStr(TotalForms), "Sticky Note"
            Forms(TotalForms - 1).Visible = False
            Forms(TotalForms - 1).lstMenu.Selected(0) = Split(MyData(i), "¬")(0)
            Forms(TotalForms - 1).lstMenu.Selected(1) = Split(MyData(i), "¬")(1)
            Forms(TotalForms - 1).lstMenu.Selected(2) = Split(MyData(i), "¬")(2)
            Forms(TotalForms - 1).txtNote.Text = Split(MyData(i), "¬")(3)
            Forms(TotalForms - 1).myDate = Split(MyData(i), "¬")(4)
            Forms(TotalForms - 1).myTime = Split(MyData(i), "¬")(5)
            Forms(TotalForms - 1).Visible = Split(MyData(i), "¬")(6)
        Next i
        Close #iFileNumber
    End If
    
    'system wide hotkey
    HotKey.CreateHotKey Me.hWnd, MOD_WIN, vbKeyS
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errorH
  Dim rtn As Long

  'this procedure receives the callbacks from the
  'System Tray icon and pops up the menu if the right
  'button is clicked.  I left the other button options
  'there, incase you want to have other options...

    'just incase we are trying to process something, let's block this out...
    'the value of X will vary depending on the scalemode setting
    If Me.ScaleMode = vbPixels Then
        rtn = X
      Else
        rtn = X / Screen.TwipsPerPixelX
    End If

    Select Case rtn
      Case WM_LBUTTONDOWN         '= &H201 - Left Button down
        'nothing happens, yet
      Case WM_LBUTTONUP           '= &H202 - Left Button up
        'nothing happens, yet
      Case WM_LBUTTONDBLCLK       '= &H203 - Left Double-click
        'nothing happens, yet
      Case WM_RBUTTONDOWN         '= &H204 - Right Button down
        'nothing happens, yet
      Case WM_RBUTTONUP           '= &H205 - Right Button up
        SetForegroundWindow Me.hWnd
        On Error Resume Next
        
        If Forms.Count >= 2 Then
            
            mnuSN(0).Enabled = True
            mnuSN(0).Caption = Forms(1).Caption
            
            For i = Forms.Count To mnuSN.UBound
                Unload mnuSN(i)
            Next i
            
            For i = 3 To Forms.Count
                Load mnuSN(i)
                mnuSN(i).Caption = Forms(i - 1).Caption
            Next i
            
        Else
            For i = Forms.Count To mnuSN.LBound
                Unload mnuSN(i)
            Next i
            mnuSN(0).Enabled = False
        End If
        
        Me.PopupMenu Me.mnuOptions
      Case WM_RBUTTONDBLCLK       '= &H206 - Right Double-click
        'nothing happens, yet
    End Select

Exit Sub

errorH:
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HotKey.DeleteHotKey Me.hWnd
    Call RemoveSystray
    For i = 1 To Forms.Count - 1
        Unload Forms(i)
    Next i
End Sub



Public Sub mnuNew_Click()

    Dim TotalForms As Integer
    ' We Fill TotalForms with the Total Number of forms +1
    TotalForms = Forms.Count + 1
    ' we call our subroutine to create the from.
    ' Totalforms is converted to a string and used for in name and caption
    Dform.CreateForm "Form" + CStr(TotalForms), "Sticky Note"
    Forms(TotalForms - 1).Caption = "Sticky Note"
    Forms(TotalForms - 1).txtTitle = "Sticky Note"
    
End Sub

Private Sub mnuExit_Click()
    
    frmMain.Show
    
    If Forms.Count > 1 Then
        a = MsgBox("You still have some Sticky Notes active (may be hidden). Do you want to save the data they contain? Next time you start Sticky Notes, these Notes will be loaded with the options as they are currently selected.", vbQuestion + vbYesNo, "Save Sticky Notes")
        If a = vbYes Then
            For i = 2 To Forms.Count
                txtSave.Text = txtSave.Text & Forms(i - 1).lstMenu.Selected(0) & "¬" & Forms(i - 1).lstMenu.Selected(1) & "¬" & Forms(i - 1).lstMenu.Selected(2) & "¬" & Forms(i - 1).txtNote.Text & "¬" & Forms(i - 1).myDate & "¬" & Forms(i - 1).myTime & "¬" & Forms(i - 1).Visible & "#~!@" & vbCrLf
            Next i
            Open App.Path & "\SND.osn" For Output As #1
                Print #1, txtSave.Text
            Close #1
        Else
            'no point having the file!
            On Error Resume Next
            Kill App.Path & "\SND.osn"
        End If
    Else
        'no sticky notes active, so don't save anything!
        On Error Resume Next
        Kill App.Path & "\SND.osn"
    End If
    
    Unload Me

End Sub

Private Sub mnuShowAll_Click()
    For i = 1 To Forms.Count - 1
        Forms(i).Visible = True
        Forms(i).WindowState = vbNormal
    Next i
End Sub

Private Sub mnuSN_Click(Index As Integer)
    
    If Index <> 0 Then
        For i = 2 To Forms.Count - 1
            'MsgBox i & " " & Index
            If Forms(i).Caption = mnuSN(Index).Caption And (Index - 1) = i Then
                Forms(i).Visible = True
                Forms(i).WindowState = vbNormal
                rtn = SetWindowPos(Forms(i).hWnd, -1, 0, 0, 0, 0, 3)
                rtn = SetWindowPos(Forms(i).hWnd, -2, 0, 0, 0, 0, 3)
                Exit For
            End If
        Next i
    Else
        Forms(1).Visible = True
        Forms(1).WindowState = vbNormal
        rtn = SetWindowPos(Forms(1).hWnd, -1, 0, 0, 0, 0, 3)
        rtn = SetWindowPos(Forms(1).hWnd, -2, 0, 0, 0, 0, 3)
    End If
End Sub
