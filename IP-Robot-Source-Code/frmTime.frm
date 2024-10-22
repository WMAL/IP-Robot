VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmTime 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                                   Data Time"
   ClientHeight    =   3360
   ClientLeft      =   12300
   ClientTop       =   5820
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjIProbot.lvButtons_H cmdOk 
      Default         =   -1  'True
      Height          =   330
      Left            =   990
      TabIndex        =   1
      Top             =   2940
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      Caption         =   "&ok"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   178
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
   Begin MSACAL.Calendar calDate 
      Height          =   2580
      Left            =   52
      TabIndex        =   0
      Top             =   255
      Width           =   3930
      _Version        =   524288
      _ExtentX        =   6932
      _ExtentY        =   4551
      _StockProps     =   1
      BackColor       =   14737632
      Year            =   2007
      Month           =   5
      Day             =   27
      DayLength       =   3
      MonthLength     =   3
      DayFontColor    =   192
      FirstDay        =   6
      GridCellEffect  =   0
      GridFontColor   =   10485760
      GridLinesColor  =   0
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   -1  'True
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjIProbot.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   330
      Left            =   60
      TabIndex        =   3
      Top             =   2940
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   582
      Caption         =   "&cancel"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   178
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   105
      TabIndex        =   2
      Top             =   45
      Width           =   3840
   End
End
Attribute VB_Name = "frmTime"
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

Private Sub calDate_DblClick()
    Call cmdOk_Click
End Sub

Private Sub cmdOk_Click()
    Select Case TheBox
        Case 0:
            If IsDate(frmRobot.comTo.Text) Then
                If DateValue(frmRobot.comTo) < DateValue(calDate.Value) Then
                    MsgBox "The chosen 'To' date appears to be earlier than 'From' date chosen here." & _
                        vbCrLf & "The 'From' date must be earlier.", vbCritical, "IP ROBOT"
                    Exit Sub
                End If
            End If
            frmRobot.txtFrom.Text = calDate.Value
        Case 1:
            If IsDate(frmRobot.txtFrom.Text) Then
                If DateValue(frmRobot.txtFrom) > DateValue(calDate.Value) Then
                    MsgBox "The chosen 'From' date appears to be later than 'To' date chosen here." & _
                        vbCrLf & "The 'To' date must be later.", vbCritical, "IP ROBOT"
                    Exit Sub
                End If
            End If
            frmRobot.comTo.Clear
            frmRobot.comTo.AddItem (calDate.Value)
            frmRobot.comTo.AddItem ("Forever")
            frmRobot.comTo.Text = frmRobot.comTo.List(0)
    End Select
    
    Unload Me
End Sub

Private Sub Form_Load()
    calDate.Value = Date
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
