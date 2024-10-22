VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                                      Options"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AD4416&
      Height          =   855
      Left            =   90
      TabIndex        =   6
      Top             =   3270
      Width           =   3885
      Begin prjIProbot.lvButtons_H cmdViewDays 
         Height          =   300
         Left            =   150
         TabIndex        =   7
         Top             =   360
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   529
         Caption         =   "days history"
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
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin prjIProbot.lvButtons_H cmdViewHours 
         Height          =   300
         Left            =   2220
         TabIndex        =   8
         Top             =   360
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   529
         Caption         =   "hours history"
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
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
   End
   Begin prjIProbot.lvButtons_H cmdClose 
      Height          =   360
      Left            =   90
      TabIndex        =   0
      Top             =   4320
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   635
      Caption         =   "c&lose"
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Log files"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AD4416&
      Height          =   2235
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   3885
      Begin VB.TextBox txtDaysClear 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "60"
         Top             =   240
         Width           =   435
      End
      Begin prjIProbot.lvButtons_H cmdClearLogs 
         Height          =   300
         Left            =   150
         TabIndex        =   3
         Top             =   1800
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         Caption         =   "&clear now"
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
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Auto-clear logs files in"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   300
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   $"frmOptions.frx":0000
         ForeColor       =   &H000000C0&
         Height          =   795
         Left            =   180
         TabIndex        =   9
         Top             =   900
         Width           =   3600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "days (1 to 240 days)"
         Height          =   195
         Left            =   2280
         TabIndex        =   5
         Top             =   300
         Width           =   1425
      End
      Begin VB.Label lblDaysClear 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Logs will be cleared next in: %DAYCLEAR%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   600
         Width           =   3630
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Converted data"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00AD4416&
      Height          =   945
      Left            =   90
      TabIndex        =   11
      Top             =   2310
      Width           =   3885
      Begin VB.TextBox txtPercentage 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1740
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "3"
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "is left (1 to 5 percent)"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   540
         Width           =   1485
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "percent of disk space"
         Height          =   255
         Left            =   2250
         TabIndex        =   14
         Top             =   300
         Width           =   1545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Delete old .PST files if"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmOptions"
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

Dim DayClear As Byte

Private Sub cmdClearLogs_Click()
    On Error Resume Next

    Dim Resp As VbMsgBoxResult
    Resp = MsgBox("This will clear the activity logs and histories." & vbCrLf & "Are you sure?", vbExclamation + vbYesNo + vbDefaultButton2, "IP ROBOT")
    If Resp = vbYes Then
        Kill (strAppPath & DayLog)
        Kill (strAppPath & DayHist)
        Kill (strAppPath & HourLog)
        Kill (strAppPath & HourHist)
        
        If Dir(strAppPath & DayHist) = "" Then Me.cmdViewDays.Enabled = False
        If Dir(strAppPath & HourHist) = "" Then Me.cmdViewHours.Enabled = False
    End If
    
    cmdClearLogs.Enabled = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdViewDays_Click()
    On Error Resume Next
    vLaunch strAppPath & DayHist
End Sub

Private Sub cmdViewHours_Click()
    On Error Resume Next
    vLaunch strAppPath & HourHist
End Sub

Private Sub Form_Load()
    Dim Temp As String
    
    If Not boStopTimer Then Me.cmdClearLogs.Enabled = False
    
    strLogLife = GetSetting("IP ROBOT", "Options", "Log life")
    CriticalFreeSpace = GetSetting("IP ROBOT", "Options", "Free disk percentage")
    
    Temp = LogClearDate
    
    If strLogLife = "" Then
        Me.txtDaysClear.Text = "60"
    Else
        Me.txtDaysClear.Text = strLogLife
    End If
    If CriticalFreeSpace = "" Then
        Me.txtPercentage.Text = "5"
        SaveSetting "IP ROBOT", "Options", "Free disk percentage", "5"
    Else
        Me.txtPercentage.Text = CriticalFreeSpace
    End If
    
    LogClearDate = Temp
    lblDaysClear.Caption = "Logs will be cleared next in: " & LogClearDate
    SaveSetting "IP ROBOT", "Options", "Clear date", LogClearDate
    
    If Dir(strAppPath & DayHist) = "" Then Me.cmdViewDays.Enabled = False
    If Dir(strAppPath & HourHist) = "" Then Me.cmdViewHours.Enabled = False

    LogClearDate = DateAdd("d", Me.txtDaysClear.Text, Date)
    lblDaysClear.Caption = Replace(lblDaysClear.Caption, "%DAYCLEAR%", LogClearDate)
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtDaysClear_Change()
    If IsNumeric(txtDaysClear.Text) Then
        If txtDaysClear.Text > 0 And txtDaysClear.Text <= 240 Then
            LogClearDate = DateAdd("d", Me.txtDaysClear.Text, Date)
            strLogLife = txtDaysClear.Text
            lblDaysClear.Caption = "Logs will be cleared next in: " & LogClearDate
            SaveSetting "IP ROBOT", "Options", "Log life", strLogLife
            SaveSetting "IP ROBOT", "Options", "Clear date", LogClearDate
            Exit Sub
        End If
    End If
    Beep
    txtDaysClear.Text = DayClear
    txtDaysClear.SelLength = Len(txtDaysClear.Text)
End Sub

Private Sub txtDaysClear_GotFocus()
    DayClear = txtDaysClear.Text
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

Private Sub txtPercentage_Change()
    If IsNumeric(txtPercentage.Text) Then
        If txtPercentage.Text > 0 And txtPercentage.Text < 6 Then
            CriticalFreeSpace = txtPercentage.Text
            SaveSetting "IP ROBOT", "Options", "Free disk percentage", CriticalFreeSpace
            Exit Sub
        End If
    End If
    Beep
    txtPercentage.Text = CriticalFreeSpace
    txtPercentage.SelLength = Len(txtPercentage.Text)
End Sub
