VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEncryptor 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                             Address Editor"
   ClientHeight    =   3870
   ClientLeft      =   12300
   ClientTop       =   5955
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjIProbot.lvButtons_H cmdAdd 
      Default         =   -1  'True
      Height          =   300
      Left            =   3390
      TabIndex        =   2
      Top             =   3435
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   529
      Caption         =   "&add"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   0
      cFHover         =   0
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   1564
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtAdd 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   124
      TabIndex        =   1
      Top             =   3435
      Width           =   3210
   End
   Begin VB.ListBox lstData 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3810
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuImport 
         Caption         =   "&Import addresses from file"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export addresses to file"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmEncryptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!
'Please read before making use of this code!
'Disclaimer: This is illegal if executed on real victims and could land you in prison for sure.
'This is intended for educational purposes only. We take no responsibility at all for your actions.
'This code is provided by EEEDS Eagle Eye Digital Security (Oman) for education purpose only.
'For more educational source codes please visit us http://www.digi77.com
'Author of this code W. Al Maawali Founder of  Eagle Eye Digital Solutions and Oman0.net can be reached via warith@digi77.com .

'Sharing knowledge is not about giving people something, or getting something from them.
'That is only valid for information sharing.
'Sharing knowledge occurs when people are genuinely interested in helping one another develop new capacities for action;
'it is about creating learning processes.
'Peter Senge
'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!

Option Explicit
Dim strSourceFile As String

Private Sub cmdAdd_Click()
    Dim strRawText As String
    Dim strAddress() As String
    Dim strAdd As String
    Dim i As Byte
    
    
    strAdd = Me.txtAdd.Text
    
    If strAdd = "" Then Exit Sub
    
    'Encrypting entry before writting it to file:
    Open (strSourceFile) For Append As 1
        Print #1, EncryptIt(strAdd)
    Close #1
    
    'Re-loading the addresses list:
    Open (strSourceFile) For Input As 1
        strRawText = Input(FileLen(strSourceFile), 1)
    Close #1
    strAddress = Split(strRawText, vbCrLf)
    Me.lstData.Clear
    'Decrypt then show addresses:
    For i = 0 To UBound(strAddress) - 1
        Me.lstData.AddItem DecryptIt(strAddress(i))
    Next i
    txtAdd.Text = ""
End Sub

Private Sub Form_Load()
    Dim strRawText As String
    Dim strAddress() As String
    Dim i As Byte
    
    strSourceFile = frmRobot.txtFilter.Text
    
    If strSourceFile <> "" Then
        If Dir(strSourceFile) <> "" Then
            Open strSourceFile For Input As 1
                strRawText = Input(FileLen(strSourceFile), 1)
            Close #1
            
            If Len(strRawText) = 0 Then Exit Sub
            
            strAddress = Split(strRawText, vbCrLf)
            
            Me.lstData.Clear
            For i = 0 To UBound(strAddress) - 1
                If strAddress(i) <> "" Then Me.lstData.AddItem DecryptIt(strAddress(i))
            Next i
        End If
    End If
End Sub
'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!
'Please read before making use of this code!
'Disclaimer: This is illegal if executed on real victims and could land you in prison for sure.
'This is intended for educational purposes only. We take no responsibility at all for your actions.
'This code is provided by EEEDS Eagle Eye Digital Security (Oman) for education purpose only.
'For more educational source codes please visit us http://www.digi77.com
'Author of this code W. Al Maawali Founder of  Eagle Eye Digital Solutions and Oman0.net can be reached via warith@digi77.com .

'Sharing knowledge is not about giving people something, or getting something from them.
'That is only valid for information sharing.
'Sharing knowledge occurs when people are genuinely interested in helping one another develop new capacities for action;
'it is about creating learning processes.
'Peter Senge
'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!

Private Sub lstData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then Call mnuRemove_Click
End Sub

Private Sub lstData_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If lstData.ListIndex = -1 Then
            mnuRemove.Enabled = False
            mnuCopy.Enabled = False
        Else
            mnuRemove.Enabled = True
            mnuCopy.Enabled = True
        End If
        PopupMenu mnuMenu
    End If
End Sub

Private Sub mnuCopy_Click()
    Clipboard.Clear
    Clipboard.SetText lstData.List(lstData.ListIndex)
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuExport_Click()
    Dim i As Integer
    
    'Initializing dialog box:
    comDlg.DialogTitle = "Open address file"
    comDlg.Filter = "Text files (*.txt)|*.txt"
    comDlg.ShowSave
    
    'File name can not be blank:
    If comDlg.FileName <> "" Then
        Open comDlg.FileName For Output As 1
            For i = 0 To lstData.ListCount - 1
                If lstData.List(i) <> "" Then Print #1, lstData.List(i)
            Next i
        Close #1
        
    MsgBox "Exporting complete!", vbInformation, "IP ROBOT"
    End If

End Sub

Private Sub mnuImport_Click()
    On Error GoTo ErrHandeler
    
    Dim Resp As VbMsgBoxResult
    Dim strRawText As String
    Dim strAddress() As String
    Dim i As Integer
    
    'Initializing dialog box:
    comDlg.DialogTitle = "Open address file"
    comDlg.Filter = "Text files (*.txt)|*.txt"
    comDlg.ShowOpen
    
    'File name can not be blank:
    If Dir(comDlg.FileName) = "" Then
        Exit Sub
    Else
        Resp = MsgBox("This will overwrite the current address list." & vbCrLf & "Are you sure?", vbExclamation + vbYesNo + vbDefaultButton2, "IP ROBOT")
        
        If Resp = vbNo Then Exit Sub
        Open (comDlg.FileName) For Input As 1
            strRawText = Input(FileLen(comDlg.FileName), 1)
        Close #1
        
        strAddress = Split(strRawText, vbCrLf)
        
        Open strSourceFile For Output As 1
            For i = 0 To UBound(strAddress)
                Print #1, EncryptIt(strAddress(i))
            Next i
        Close #1
         
        Call Form_Load
    End If
    
    Exit Sub
ErrHandeler:
    MsgBox Err.Description, , "IP ROBOT"
End Sub

Private Sub mnuRemove_Click()
    Dim strRawText As String
    Dim strAddress() As String
    Dim i As Byte
    
    Me.lstData.RemoveItem Me.lstData.ListIndex
    
    'Save changes:
    Open (strSourceFile) For Output As 1
        If lstData.ListCount = 0 Then
            Close #1
            Exit Sub
        Else
            For i = 0 To lstData.ListCount - 1
                Print #1, EncryptIt(lstData.List(i))
            Next i
            Close #1
        End If
    
    Open (strSourceFile) For Input As 1
        strRawText = Input(FileLen(strSourceFile), 1)
    Close #1

    strAddress = Split(strRawText, vbCrLf)
    
    Me.lstData.Clear
    For i = 0 To UBound(strAddress) - 1
        Me.lstData.AddItem DecryptIt(strAddress(i))
    Next i
End Sub

'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!
'Please read before making use of this code!
'Disclaimer: This is illegal if executed on real victims and could land you in prison for sure.
'This is intended for educational purposes only. We take no responsibility at all for your actions.
'This code is provided by EEEDS Eagle Eye Digital Security (Oman) for education purpose only.
'For more educational source codes please visit us http://www.digi77.com
'Author of this code W. Al Maawali Founder of  Eagle Eye Digital Solutions and Oman0.net can be reached via warith@digi77.com .

'Sharing knowledge is not about giving people something, or getting something from them.
'That is only valid for information sharing.
'Sharing knowledge occurs when people are genuinely interested in helping one another develop new capacities for action;
'it is about creating learning processes.
'Peter Senge
'!!!!!***************!!!!!!!!!******************!!!!!!!!!!!!**********!

