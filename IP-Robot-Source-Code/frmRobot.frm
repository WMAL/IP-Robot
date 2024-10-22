VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRobot 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mail Robot - version 1.05 - by Dr Jeeni"
   ClientHeight    =   7755
   ClientLeft      =   6825
   ClientTop       =   3270
   ClientWidth     =   9240
   Icon            =   "frmRobot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmRobot.frx":030A
   ScaleHeight     =   7755
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrFreeSpacer 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   8610
      Top             =   2985
   End
   Begin prjIProbot.lvButtons_H cmdAction 
      Default         =   -1  'True
      Height          =   495
      Left            =   115
      TabIndex        =   31
      ToolTipText     =   "Click here to start processing data according on your choices."
      Top             =   7140
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   873
      Caption         =   "&START"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFHover         =   -2147483640
      cBhover         =   -2147483633
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   14215660
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   7635
      Top             =   2925
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrChecker 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8130
      Top             =   2985
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Source and Destination Folders"
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
      Height          =   2655
      Left            =   115
      TabIndex        =   32
      Top             =   90
      Width           =   4470
      Begin VB.TextBox txtOtherPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "path not defined"
         Top             =   2160
         Width           =   2925
      End
      Begin VB.TextBox txtWebPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "path not defined"
         Top             =   1620
         Width           =   2925
      End
      Begin VB.TextBox txtPOPpath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "path not defined"
         Top             =   1080
         Width           =   2925
      End
      Begin VB.TextBox txtFilteredPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "path not defined"
         Top             =   540
         Width           =   2925
      End
      Begin prjIProbot.lvButtons_H cmdOthers 
         Height          =   300
         Left            =   3195
         TabIndex        =   14
         ToolTipText     =   "Click to define data folder path."
         Top             =   2145
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         Caption         =   "..."
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin prjIProbot.lvButtons_H cmdWEBMAIL 
         Height          =   300
         Left            =   3195
         TabIndex        =   10
         ToolTipText     =   "Click to define Webmail data folder path."
         Top             =   1605
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         Caption         =   "..."
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin prjIProbot.lvButtons_H cmdPOPSMTP 
         Height          =   300
         Left            =   3195
         TabIndex        =   6
         ToolTipText     =   "Click to define POP and SMTP data folder path."
         Top             =   1065
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         Caption         =   "..."
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin VB.OptionButton optOthers 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   ".eml files source (Data4)"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   225
         TabIndex        =   12
         ToolTipText     =   "Select this to defined path of data other than the three options above."
         Top             =   1935
         Width           =   2055
      End
      Begin VB.OptionButton optFiltered 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   ".eml files source (Data1)"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   225
         TabIndex        =   0
         ToolTipText     =   "Select this to convert data brought from Data1 folder (path definition required)."
         Top             =   315
         Width           =   2055
      End
      Begin VB.OptionButton optPopSmtp 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   ".eml files source (Data2)"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   225
         TabIndex        =   4
         ToolTipText     =   "Select this to convert POP and SMTP data brought from Data2 folder (path definition required)."
         Top             =   855
         Width           =   2055
      End
      Begin VB.OptionButton optWebMail 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   ".eml files source (Data3)"
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   225
         TabIndex        =   8
         ToolTipText     =   "Select this to convert webmail data brought from Data3 folder (path definition required)."
         Top             =   1395
         Width           =   2055
      End
      Begin prjIProbot.lvButtons_H cmdFiltered 
         Height          =   300
         Left            =   3195
         TabIndex        =   2
         ToolTipText     =   "Click to define Filtered data folder path."
         Top             =   525
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         Caption         =   "..."
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin prjIProbot.lvButtons_H cmdOpenFiltered 
         Height          =   300
         Left            =   3540
         TabIndex        =   3
         ToolTipText     =   "Click to open Filtered data folder."
         Top             =   525
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
         Caption         =   "open"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   1
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
         cBhover         =   -2147483633
         LockHover       =   1
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin prjIProbot.lvButtons_H cmdOpenOthers 
         Height          =   300
         Left            =   3540
         TabIndex        =   15
         ToolTipText     =   "Click to open data folder."
         Top             =   2145
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
         Caption         =   "open"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   1
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
      Begin prjIProbot.lvButtons_H cmdOpenWeb 
         Height          =   300
         Left            =   3540
         TabIndex        =   11
         ToolTipText     =   "Click to open Webmail data folder."
         Top             =   1605
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
         Caption         =   "open"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   1
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
      Begin prjIProbot.lvButtons_H cmdOpenPOP 
         Height          =   300
         Left            =   3540
         TabIndex        =   7
         ToolTipText     =   "Click to open POP and SMTP data folder."
         Top             =   1065
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
         Caption         =   "open"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   1
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   915
      Left            =   115
      TabIndex        =   33
      Top             =   2610
      Width           =   4470
      Begin VB.TextBox txtDest 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "destination path goes here"
         ToolTipText     =   "The path of the destination folder. The folder in which resulted .PST files will be saved."
         Top             =   405
         Width           =   2925
      End
      Begin prjIProbot.lvButtons_H cmdDestination 
         Height          =   300
         Left            =   3195
         TabIndex        =   17
         ToolTipText     =   "Click to define the destination folder path."
         Top             =   390
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         Caption         =   "..."
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
      Begin prjIProbot.lvButtons_H cmdOpenDest 
         Height          =   300
         Left            =   3540
         TabIndex        =   18
         ToolTipText     =   "Click to open the destination folder."
         Top             =   390
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
         Caption         =   "open"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   1
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
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Destination folder (pst output):"
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Other Options"
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
      Height          =   2025
      Left            =   4664
      TabIndex        =   36
      Top             =   1500
      Width           =   4470
      Begin VB.TextBox txtFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "filter file goes here"
         Top             =   600
         Width           =   2925
      End
      Begin VB.CheckBox chkFilterON 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enable &filter "
         Height          =   195
         Left            =   210
         TabIndex        =   23
         ToolTipText     =   "Allows using a list of targets e-mail addresses to filter converted data."
         Top             =   330
         Width           =   1215
      End
      Begin VB.ComboBox comHours 
         Height          =   315
         ItemData        =   "frmRobot.frx":41BE
         Left            =   690
         List            =   "frmRobot.frx":41DA
         Style           =   2  'Dropdown List
         TabIndex        =   27
         ToolTipText     =   $"frmRobot.frx":41F8
         Top             =   1050
         Width           =   675
      End
      Begin prjIProbot.lvButtons_H cmdFilter 
         Height          =   300
         Left            =   3180
         TabIndex        =   25
         ToolTipText     =   "Define the file that stores a list of targets e-mail addresses. The addresses will be used to filter data."
         Top             =   585
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         Caption         =   "..."
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
      Begin prjIProbot.lvButtons_H cmdEdit 
         Height          =   300
         Left            =   3540
         TabIndex        =   26
         ToolTipText     =   "Edit the important e-mail addresses list."
         Top             =   585
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   529
         Caption         =   "&edit"
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   1
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
      Begin prjIProbot.lvButtons_H cmdOptions 
         Height          =   360
         Left            =   180
         TabIndex        =   28
         ToolTipText     =   "Display options to control the way ROBOT will be handing log files."
         Top             =   1500
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
         Caption         =   "more options"
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
         cFHover         =   0
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fuse"
         Height          =   195
         Left            =   180
         TabIndex        =   38
         Top             =   1170
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "hours into each .pst file"
         Height          =   195
         Left            =   1470
         TabIndex        =   37
         Top             =   1170
         Width           =   1635
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Conversion Timing"
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
      Height          =   1395
      Left            =   4664
      TabIndex        =   35
      Top             =   90
      Width           =   4470
      Begin VB.TextBox txtFrom 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "date not defined"
         Top             =   330
         Width           =   3015
      End
      Begin VB.ComboBox comTo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmRobot.frx":428E
         Left            =   600
         List            =   "frmRobot.frx":4295
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   870
         Width           =   3015
      End
      Begin prjIProbot.lvButtons_H cmdFrom 
         Height          =   300
         Left            =   3660
         TabIndex        =   20
         ToolTipText     =   "Click here to set the 'From' date  of the data you want to process."
         Top             =   330
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   529
         Caption         =   "set"
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
      Begin prjIProbot.lvButtons_H cmdTo 
         Height          =   300
         Left            =   3660
         TabIndex        =   22
         ToolTipText     =   "Click here to set the 'To' date  of the data you want to process. Choose 'Forever' for a continuous data process."
         Top             =   885
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   529
         Caption         =   "set"
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
         Caption         =   "From:"
         Height          =   195
         Left            =   180
         TabIndex        =   40
         Top             =   435
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "To:"
         Height          =   195
         Left            =   330
         TabIndex        =   39
         Top             =   990
         Width           =   240
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Activity Log"
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
      Height          =   3495
      Left            =   115
      TabIndex        =   41
      Top             =   3540
      Width           =   9015
      Begin VB.TextBox txtHoursLog 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1350
         Left            =   150
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   1980
         Width           =   8730
      End
      Begin VB.TextBox txtDaysLog 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1350
         Left            =   150
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   450
         Width           =   8730
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HOURS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   150
         TabIndex        =   43
         Top             =   1770
         Width           =   8730
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DAYS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   150
         TabIndex        =   42
         Top             =   240
         Width           =   8730
      End
   End
   Begin prjIProbot.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   8490
      TabIndex        =   45
      ToolTipText     =   "Click here to start processing data according on your choices."
      Top             =   7140
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   873
      Caption         =   "&close"
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
      cFHover         =   -2147483640
      cBhover         =   -2147483633
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   14215660
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1890
      TabIndex        =   44
      Top             =   7290
      Width           =   6525
   End
End
Attribute VB_Name = "frmRobot"
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

Dim arrFiles(1 To 24) As String     'Stores the names of cereated batch files, so they can be deleted by form_unload event
Dim intFileLen, intFinished As Integer 'intFileLen, stores the length of the hours log file, checked by the timer, size change forces log update
                                       'intFinished, counts the number of days processed.
Dim strDest, strSource, strAddressFile, strFilter, strDataType, strTemp As String
Dim strYear, strMonth, strDay, strHour, strHours As String
Dim EndDate, DayBegin As Date
Dim FilesCount, LinesCount As Byte 'These are used to determine the number of list files to be created by a loop at cmdAction event
                                   'and the number of lines in each file.

Private Sub chkFilterON_Click()
    SaveSetting "Mail Robot", "Options", "Filter", chkFilterON.Value
    
    If chkFilterON.Value = 0 Then
        Me.txtFilter.Enabled = False
        Me.cmdFilter.Enabled = False
        Me.cmdEdit.Enabled = False
    Else
        Me.txtFilter.Enabled = True
        Me.cmdFilter.Enabled = True
        Me.cmdEdit.Enabled = True
    End If
End Sub


Private Sub cmdFilter_Click()

    'Initializing dialog box:
    comDlg.DialogTitle = "Open address file"
    comDlg.Filter = "dll (*.dll)|*.dll"
    comDlg.ShowOpen
    strFilter = comDlg.FileName
    
    'File name can not be blank:
    If comDlg.FileName = "" Then Exit Sub
    
    'Save to reg:
    SaveSetting "Mail Robot", "Paths", "Filter", strFilter
    txtFilter.Text = strFilter
End Sub

Private Sub cmdDestination_Click()
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        'sBuffer value is the directory that the user choose from the dialog.
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        If Right(sBuffer, 1) <> "\" Then sBuffer = sBuffer & "\"
    End If
    
    If sBuffer <> "\" And sBuffer <> "" Then
        txtDest.Text = sBuffer
        strDest = txtDest.Text
        SaveSetting "Mail Robot", "Paths", "Destination", sBuffer
    End If
End Sub

Private Sub cmdAction_Click()
    Dim i, j, btIndex As Byte
    Dim strFileName As String
    Dim Resp As VbMsgBoxResult
    Dim strNofDays As String
    
    'Reset the source folder:
    strSource = ""
    
    If cmdAction.Caption = "&START" Then
    
        'Setting the data source:
        If Me.optFiltered Then
            strSource = Me.txtFilteredPath
            strDataType = "FILTERED"
        ElseIf Me.optOthers Then
            strSource = Me.txtOtherPath
            strDataType = "OTHERS"
        ElseIf Me.optPopSmtp Then
            strSource = Me.txtPOPpath
            strDataType = "POPSMTP"
        ElseIf Me.optWebMail Then
            strSource = Me.txtWebPath
            strDataType = "WEBMAIL"
        End If
        
        'Validating the source, destination and filter file:
        If strSource = "" Or Dir(strSource, vbDirectory) = "" Or Dir(txtDest.Text, vbDirectory) = "" Then
            MsgBox "Invalid source or destination folder selected. Please correct.", vbCritical, "Mail Robot"
            GoTo ReEnable
            Exit Sub
        ElseIf chkFilterON.Value = 1 Then
            If Dir(txtFilter.Text) = "" Then
                MsgBox "Could not open filter file or filter file does not exist.", vbCritical, "Mail Robot"
                GoTo ReEnable
                Exit Sub
            End If
        End If
        
        'Resetting the finished days:
        intFinished = 0
        
        'Confirmation:
        Resp = MsgBox("Start conversion process, are you sure?", vbQuestion + vbYesNo, "Mail Robot")
        If Resp = vbNo Then Exit Sub
        
        'Disabling Controls:
        Me.Frame1.Enabled = False
        Me.Frame2.Enabled = False
        Me.Frame3.Enabled = False
        Me.chkFilterON.Enabled = False
        Me.txtFilter.Enabled = False
        Me.cmdFilter.Enabled = False
        Me.cmdEdit.Enabled = False
        Me.comHours.Enabled = False
        
        'Setting MyNow:
        MyNow = CDate(txtFrom.Text) & " 00:00:01"
        
        'Setting the end date, and number of days to be processed:
        If IsDate(comTo.Text) Then
            EndDate = CDate(comTo.Text)
            strNofDays = DateDiff("d", txtFrom.Text, comTo.Text) + 1
        Else
            EndDate = "31/12/9999" ' The max date a calendar can take! DOOM DAY may happen before this!
            strNofDays = "UNLIMITED"
        End If
        
        'NEW: Creating log and history files headers:
        Open (strAppPath & DayLog) For Output As 1
                Print #1, "---------------------------------------------------------------------------------------------"; vbCrLf; _
                          "Mail Robot DAYS ACTIVITY LOG - DATE: "; Format(Now, "dd/mm/yyyy HH:nn:ss"); vbCrLf; _
                          "---------------------------------------------------------------------------------------------"; vbCrLf; _
                          "START DATE:      ["; txtFrom.Text; "]"; vbCrLf; _
                          "END DATE:        ["; UCase(comTo.Text); "]"; vbCrLf; _
                          "DAYS TO PROCESS: ["; strNofDays; "]"; vbTab; vbCrLf; _
                          "DAYS PROCESSED:  [0]"; vbCrLf; _
                          "DATA TYPE:       ["; strDataType; "]"; vbCrLf; _
                          "---------------------------------------------------------------------------------------------"; vbCrLf; _
                          "DATA DATE      STARTED             FINISHED            DURATION"; vbCrLf; _
                          "---------      --------------      --------------      --------"
        Close #1

        
        If Dir(strAppPath & DayHist) = "" Then
            Open (strAppPath & DayHist) For Output As 1
                Print #1, "---------------------------------------------------------------"; vbCrLf; _
                          "Mail Robot DAYS ACTIVITY HISTORY"; vbCrLf; _
                          "---------------------------------------------------------------"; vbCrLf; _
                          "DATA DATE      STARTED             FINISHED            DURATION"; vbCrLf; _
                          "---------      --------------      --------------      --------"
            Close #1
        End If
        
        Open (strAppPath & HourLog) For Output As 1
                Print #1, "---------------------------------------------------------------------------------------------"; vbCrLf; _
                          "Mail Robot HOURS ACTIVITY LOG - DATE: "; Format(Now, "dd/mm/yyyy HH:nn:ss"); vbCrLf; _
                          "---------------------------------------------------------------------------------------------"; vbCrLf; _
                          "DATA DATE          HOURS            FINISHED"; vbCrLf; _
                          "---------          -------          --------------------------"
        Close #1

        
        If Dir(strAppPath & HourHist) = "" Then
            Open (strAppPath & HourHist) For Output As 1
                Print #1, "--------------------------------------------------------------"; vbCrLf; _
                          "Mail Robot HOURS ACTIVITY HISTORY"; vbCrLf; _
                          "--------------------------------------------------------------"; vbCrLf; _
                          "DATA DATE          HOURS            FINISHED"; vbCrLf; _
                          "---------          -------          --------------------------"
            Close #1
        End If
        
        'PRINTING FIRST LINE IN THE DAYS LOG FILE:
        DayBegin = Now
        
        Open (strAppPath & DayLog) For Append As 1
            Print #1, Format(MyNow, "dd/mm/yy"); "       "; Format(DayBegin, "dd/mm/yy hh:nn"); "      IN PROGRESS         N/A"
        Close #1
        'Now, after setting the content of the days log file, we are loading it to the box:
        Open (strAppPath & DayLog) For Input As 1
            txtDaysLog.Text = Input(FileLen(strAppPath & DayLog) - 2, 1)
        Close #1
        
        'Updating the Hours Log textbox:
        Open (strAppPath & HourLog) For Input As 1
            txtHoursLog.Text = Input(FileLen(strAppPath & HourLog) - 2, 1)
        Close #1
            
        'Measuring/setting the size of hours log, which will be checked later by the timer:
        intFileLen = FileLen(strAppPath & HourLog)
        
        'Calling the lister proc, which will be creating all .lst files:
        Call TheLister
        
        boStopTimer = False
        tmrChecker.Enabled = True
        tmrFreeSpacer.Enabled = True
        
        'Change the caption of the button:
        cmdAction.Caption = "&STOP"
    Else
        'Confirmation:
        Resp = MsgBox("Stop conversion process, are you sure?", vbQuestion + vbYesNo, "Mail Robot")
        If Resp = vbNo Then Exit Sub
        
        'Stoping will only be about disabling the timer, enabling control and changing the caption:
        cmdAction.Caption = "&START"
        tmrFreeSpacer.Enabled = False
        boStopTimer = True
ReEnable:
        'Enabling controls:
        Me.Frame1.Enabled = True
        Me.Frame2.Enabled = True
        Me.Frame3.Enabled = True
        Me.chkFilterON.Enabled = True
        Me.txtFilter.Enabled = True
        Me.cmdFilter.Enabled = True
        Me.cmdEdit.Enabled = True
        Me.comHours.Enabled = True
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

Private Sub TheLister()
    Dim i, j, n, m As Byte
    Dim strFileName, strList, Hour1, Hour2 As String
    
    '1st, Create the .lst files:
    FilesCount = 24 / comHours.Text
    LinesCount = 24 / FilesCount
    
    For i = 1 To FilesCount
        For j = 1 To LinesCount
            strYear = Right(Year(DateAdd("h", n, MyNow)), 1)
            strMonth = Month(DateAdd("h", n, MyNow))
            strDay = Day(DateAdd("h", n, MyNow))
            strHour = Hour(DateAdd("h", n, MyNow))
            strList = strList & "+" & strSource & strYear & "\" & strMonth & "\" & strDay & "\" & strHour & "\*.eml" & vbCrLf
            n = n + 1
            If j = 1 Then Hour1 = strHour
            If j = LinesCount Then Hour2 = strHour
        Next j
        If Len(Hour1) = 1 Then Hour1 = "0" & Hour1
        If Len(Hour2) = 1 Then Hour2 = "0" & Hour2
        
        strHours = "(" & Hour1 & "-" & Hour2 & ")"
        
        'Choosing a random file name for the .lst:
        Randomize
        strFileName = strHours & LCase(Hex(Rnd * (Format(Now, "ddMMyyHHnnss")) / 10000))
    
        'Printing the content of list file:
        Open strTemp & strFileName & ".lst" For Output As 1
            Print #1, strList
        Close #1
        strList = ""
        
        'Storing the name of the file in an array:
        arrFiles(i) = strFileName
        
        '2nd, Decrypt the e-mail addresses filter file, if user has chosen to use it:
            If chkFilterON.Value = 1 Then
                Dim strRawText As String
                Dim strArray() As String
                
                Open strFilter For Input As 1
                    strRawText = Input(FileLen(strFilter), 1)
                Close #1
                strArray = Split(strRawText, vbCrLf)
                Randomize
                strAddressFile = Hex(Rnd * (Format(Date, "ddMMyy")) / 10)
                Open strTemp & strAddressFile For Output As 1
                    For m = 0 To UBound(strArray)
                        Print #1, DecryptIt(strArray(m))
                    Next m
                Close #1
                'Hiding the filter file:
                SetAttr strTemp & strAddressFile, vbSystem + vbHidden
            End If
        
        'Now, as we created the .lst files, call IPworks
        Call IPworx(strFileName, strDataType & strHours & ".pst")
    Next i
    

    'Finally, update MyNow (adding a single day, the day to be processed next):
    MyNow = DateAdd("d", 1, DateValue(MyNow)) & " 00:00:01"
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



Private Sub IPworx(strActionFile, strResultPST As String)
    Dim strBatch As String
    
    'Building a dynamic destination folders structure:
    If Dir(txtDest.Text & "\" & strYear, vbDirectory) = "" Then
        MkDir (txtDest.Text & "\" & strYear)
        MkDir (txtDest.Text & "\" & strYear & "\" & strMonth)
        MkDir (txtDest.Text & "\" & strYear & "\" & strMonth & "\" & strDay)
    ElseIf Dir(txtDest.Text & "\" & strYear & "\" & strMonth, vbDirectory) = "" Then
        MkDir (txtDest.Text & "\" & strYear & "\" & strMonth)
        MkDir (txtDest.Text & "\" & strYear & "\" & strMonth & "\" & strDay)
    ElseIf Dir(txtDest.Text & "\" & strYear & "\" & strMonth & "\" & strDay, vbDirectory) = "" Then
        MkDir (txtDest.Text & "\" & strYear & "\" & strMonth & "\" & strDay)
    End If
    
    'Setting Destination folder:
    strDest = txtDest.Text & strYear & "\" & strMonth & "\" & strDay & "\"
    
    'Based on the options chosen, set the content of the batch file:
    If chkFilterON.Value = 1 Then
        strResultPST = "[F]-" & strResultPST
        strBatch = "a4m.exe " & strTemp & strActionFile & ".lst " & Chr(34) & strDest & strResultPST & Chr(34) & " /s /d /Unattended /Unicodepst /Include=" & Chr(34) & strTemp & strAddressFile & Chr(34)
    Else
        strBatch = "a4m.exe " & strTemp & strActionFile & ".lst " & Chr(34) & strDest & strResultPST & Chr(34) & " /s /d /Unattended /Unicodepst"
    End If
    
    'Printing the batch file:
    Open strTemp & strActionFile & ".bat" For Output As 1
        Print #1, "@echo off"
        Print #1, strBatch
        'This will update the log file:
        Print #1, "echo " & Format(MyNow, "dd/mm/yy") & "           " & strHours & "          %date% %time% >> " & Chr(34) & strAppPath & HourLog & Chr(34)
        Print #1, "echo " & Format(MyNow, "dd/mm/yy") & "           " & strHours & "          %date% %time% >> " & Chr(34) & strAppPath & HourHist & Chr(34)
        'This will delete all files at the end of data conversion:
        Print #1, "del " & strTemp & strActionFile & ".lst"
        If chkFilterON.Value = 1 Then Print #1, "del " & strTemp & strAddressFile & "/A:h s"
        Print #1, "del " & strTemp & strActionFile & ".bat"
    Close #1
    
    'Starting the batch file:
    vLaunch strTemp & strActionFile & ".bat"
    
End Sub
Private Sub cmdEdit_Click()
    If Dir(txtFilter.Text) = "" Then
        MsgBox "Could not open filter file, or filter file does not exist.", vbCritical, "Mail Robot"
    Else
        frmEncryptor.Show vbModal
    End If
End Sub

Private Sub cmdFiltered_Click()
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        'sBuffer value is the directory that the user choose from the dialog.
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        If Right(sBuffer, 1) <> "\" Then sBuffer = sBuffer & "\"
    End If
    
    If sBuffer <> "\" And sBuffer <> "" Then
        Me.optFiltered.Enabled = True
        Me.txtFilteredPath.Text = sBuffer
        SaveSetting "Mail Robot", "Paths", "Filtered data", sBuffer
    End If
End Sub

Private Sub cmdFrom_Click()
    TheBox = TXT_FROM
    frmTime.Show vbModal
End Sub

Private Sub cmdOpenDest_Click()
    If Dir(Me.txtDest.Text, vbDirectory) <> "" Then Shell "explorer.exe " & Me.txtDest.Text, vbNormalFocus
End Sub

Private Sub cmdOpenFiltered_Click()
    If Dir(Me.txtFilteredPath.Text, vbDirectory) <> "" Then Shell "explorer.exe " & Me.txtFilteredPath.Text, vbNormalFocus
End Sub

Private Sub cmdOpenOthers_Click()
    If Dir(Me.txtOtherPath.Text, vbDirectory) <> "" Then Shell "explorer.exe " & Me.txtOtherPath.Text, vbNormalFocus
End Sub

Private Sub cmdOpenPOP_Click()
    If Dir(Me.txtPOPpath.Text, vbDirectory) <> "" Then Shell "explorer.exe " & Me.txtPOPpath.Text, vbNormalFocus
End Sub

Private Sub cmdOpenWeb_Click()
    If Dir(Me.txtWebPath.Text, vbDirectory) <> "" Then Shell "explorer.exe " & Me.txtWebPath.Text, vbNormalFocus
End Sub

Private Sub cmdOptions_Click()
    frmOptions.Show vbModal
End Sub

Private Sub cmdOthers_Click()
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        'sBuffer value is the directory that the user choose from the dialog.
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        If Right(sBuffer, 1) <> "\" Then sBuffer = sBuffer & "\"
    End If
    
    If sBuffer <> "\" And sBuffer <> "" Then
        Me.optOthers.Enabled = True
        Me.txtOtherPath.Text = sBuffer
        SaveSetting "Mail Robot", "Paths", "Others", sBuffer
    End If
End Sub

Private Sub cmdPOPSMTP_Click()
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        'sBuffer value is the directory that the user choose from the dialog.
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        If Right(sBuffer, 1) <> "\" Then sBuffer = sBuffer & "\"
    End If
    
    If sBuffer <> "\" And sBuffer <> "" Then
        Me.optPopSmtp.Enabled = True
        Me.txtPOPpath.Text = sBuffer
        SaveSetting "Mail Robot", "Paths", "POP and SMTP", sBuffer
    End If
End Sub

Private Sub cmdTo_Click()
    TheBox = TXT_TO
    frmTime.Show vbModal
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

Private Sub cmdWEBMAIL_Click()
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        'sBuffer value is the directory that the user choose from the dialog.
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        If Right(sBuffer, 1) <> "\" Then sBuffer = sBuffer & "\"
    End If
    
    If sBuffer <> "\" And sBuffer <> "" Then
        Me.optWebMail.Enabled = True
        Me.txtWebPath.Text = sBuffer
        SaveSetting "Mail Robot", "Paths", "Web mail", sBuffer
    End If
End Sub

Private Sub comHours_Click()
    SaveSetting "Mail Robot", "Options", "Number of hours", comHours.Text
End Sub

Private Sub Form_Load()
    Dim strFilterOpt As String
    Dim strNofHours As String
    Dim strActiveSource As String
    Dim strPOPpath, strWebPath, strFilteredPath, strOtherPath As String
    
    'within the Temp directory, it's better to create an exclusive folder for the ROBOT:
    strTemp = Environ("temp") & "\"
    If Dir(strTemp & "robot", vbDirectory) = "" Then
        MkDir (strTemp & "robot")
    End If
    
    '...
    strTemp = strTemp & "robot\"
    strAppPath = App.Path & "\"
    txtFrom.Text = DateAdd("d", -1, Date)
    comTo.Text = comTo.List(0)
    intFinished = 0
    
    'Loading values from the system registry:
    strDest = GetSetting("Mail Robot", "Paths", "Destination")
    strFilter = GetSetting("Mail Robot", "Paths", "Filter")
    strFilterOpt = GetSetting("Mail Robot", "Options", "Filter")
    strNofHours = GetSetting("Mail Robot", "Options", "Number of hours")
    LogClearDate = GetSetting("Mail Robot", "Options", "Log clear date")
    CriticalFreeSpace = GetSetting("Mail Robot", "Options", "Free disk percentage")
    
    'Data source:
    strPOPpath = GetSetting("Mail Robot", "Paths", "POP and SMTP")
    strWebPath = GetSetting("Mail Robot", "Paths", "Web mail")
    strFilteredPath = GetSetting("Mail Robot", "Paths", "Filtered data")
    strOtherPath = GetSetting("Mail Robot", "Paths", "Others")
    
    
    'N: Days and Hours logs and histories file names are asigned here:
    DayHist = "DH-" & Format(Now, "ddMMyyHHnnss") & ".txt"
    DayLog = Replace(DayHist, "DH", "DL")
    HourHist = Replace(DayHist, "DH", "HH")
    HourLog = Replace(DayHist, "DH", "HL")


    '....
    If CriticalFreeSpace = "" Then CriticalFreeSpace = 3
    
    If strPOPpath = "" Then
        Me.optPopSmtp.Enabled = False
    Else
        Me.txtPOPpath.Text = strPOPpath
    End If
    
    If strWebPath = "" Then
        Me.optWebMail.Enabled = False
    Else
        Me.txtWebPath.Text = strWebPath
    End If
    
    If strFilteredPath = "" Then
        Me.optFiltered.Enabled = False
    Else
        Me.txtFilteredPath.Text = strFilteredPath
    End If
    
    If strOtherPath = "" Then
        Me.optOthers.Enabled = False
    Else
        Me.txtOtherPath.Text = strOtherPath
    End If
    
    If strNofHours <> "" Then
        comHours.Text = strNofHours
    Else
        comHours.Text = "2"
    End If
    
    If strFilter <> "" Then Me.txtFilter.Text = strFilter
    If strDest <> "" Then Me.txtDest.Text = strDest
    If LogClearDate = "" Then
        LogClearDate = DateAdd("d", 60, Date)
        SaveSetting "Mail Robot", "Options", "Clear date", LogClearDate
    End If
    If strFilterOpt = "" Then
        Me.chkFilterON.Value = 0
    Else
        Me.chkFilterON.Value = Val(strFilterOpt)
    End If
    'Active data source (the data source that was chosen in Robot's last setion)
    strActiveSource = GetSetting("Mail Robot", "Options", "Active data source")
    
    Select Case strActiveSource
        Case 1: Me.optFiltered.Value = True
        Case 2: Me.optPopSmtp.Value = True
        Case 3: Me.optWebMail.Value = True
        Case 4: Me.optFiltered.Value = True
    End Select
       
    If chkFilterON.Value = 0 Then
        Me.txtFilter.Enabled = False
        Me.cmdFilter.Enabled = False
        Me.cmdEdit.Enabled = False
    Else
        Me.txtFilter.Enabled = True
        Me.cmdFilter.Enabled = True
        Me.cmdEdit.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    If cmdAction.Caption = "&START" Then GoTo DeleteFiles
    
    Dim Resp As VbMsgBoxResult
    Resp = MsgBox("This will close the ROBOT." & vbCrLf & _
                  "Any pending conversion processes will be terminated." & _
                  vbCrLf & vbCrLf & " Are you sure you want to do this?", _
                  vbQuestion + vbYesNo + vbDefaultButton2, "Mail Robot")
    If Resp = vbNo Then
        Cancel = True
    Else
DeleteFiles:
        'In case user has chosen to close the robot, all its files located in the
        ' temp folder have to be deleted. Those files' names are stored in an array
        ' by the time they were created (in the TheLister procedure).
    'So,
        Dim i As Byte
        For i = 1 To 24
            If arrFiles(i) <> "" Then
                Kill (strTemp & arrFiles(i) & ".lst")
                Kill (strTemp & arrFiles(i) & ".bat")
            End If
        Next i
    End If
End Sub

Private Sub lvButtons_H1_Click()
    Unload Me
End Sub

Private Sub optFiltered_Click()
    SaveSetting "Mail Robot", "Options", "Active data source", "1"
End Sub

Private Sub optPopSmtp_Click()
    SaveSetting "Mail Robot", "Options", "Active data source", "2"
End Sub

Private Sub optWebMail_Click()
    SaveSetting "Mail Robot", "Options", "Active data source", "3"
End Sub

Private Sub optOthers_Click()
    SaveSetting "Mail Robot", "Options", "Active data source", "4"
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

Private Sub tmrChecker_Timer()
    On Error Resume Next 'mainly put here to cancel errors when deleting log files
    
    'Updating the status textbox:
    Dim i As Byte
    Dim strDaysDiff, strHoursDiff As String
    Dim strLineToAdd As String
    
    If intFileLen <> FileLen(strAppPath & HourLog) Then
        Open (strAppPath & HourLog) For Input As 1
            txtHoursLog.Text = Input(FileLen(strAppPath & HourLog), 1)
        Close #1
        'Scrolling Down:
        txtHoursLog.SelStart = Len(txtHoursLog.Text)
        txtHoursLog.SelLength = 1
        'Setting the new file size
        intFileLen = FileLen(strAppPath & HourLog)
    End If
     
    For i = 1 To FilesCount
    'this is checking for the existence of all .lst files...
    'if any of the files is still there, then conversion is still in progress
    ' so, do not proceed:
        If Dir(strTemp & arrFiles(i) & ".lst") <> "" Then Exit Sub
    Next i
    
    '--------------
    '(N)Being after the above line, means that this is an end of all batch files for some day.
    'That day may or may not be the last day in a group of days (in case an end date was defined)
    
    'Next is to update the Days log file, since a day has just been finished.
    'But the update occures ONLY if the process is NOT being held till the target date is met:
    'So,
    If InStr(1, txtDaysLog.Text, "HELD") = 0 Then
        'Finding the duration in days and hours:
        strDaysDiff = DateDiff("d", DayBegin, Date) & "D"
        strHoursDiff = DateDiff("h", TimeValue(DayBegin), Format(Time, "HH:nn:ss"))
        strHoursDiff = 24 * ((strHoursDiff / 24) - Fix(strHoursDiff / 24)) & "H"
        'Fixing length
        If Len(strDaysDiff) = 2 Then strDaysDiff = "0" & strDaysDiff
        If Len(strHoursDiff) = 2 Then strHoursDiff = "0" & strHoursDiff
        'Replacing:
        txtDaysLog.Text = Replace(txtDaysLog.Text, "IN PROGRESS", Format(Now, "dd/MM/yy HH:mm"))
        txtDaysLog.Text = Replace(txtDaysLog.Text, "[" & intFinished & "]", "[" & intFinished + 1 & "]")
        txtDaysLog.Text = Replace(txtDaysLog.Text, "   N/A", strDaysDiff & " " & strHoursDiff)
        'Finished days counter:
        intFinished = intFinished + 1
         
        'Scrolling down:
        txtDaysLog.SelStart = Len(txtDaysLog.Text)
        txtDaysLog.SelLength = 1
        
        strLineToAdd = Right(txtDaysLog.Text, 62)
         
        'Sending the update to the days log
        Open (strAppPath & DayLog) For Output As 1
            Print #1, txtDaysLog.Text
        Close #1
        
        'And the data portion of the same file to the days history file:
        Open (strAppPath & DayHist) For Append As 1
            Print #1, strLineToAdd
        Close #1
    
        'Clearing the hours log..
        Open (strAppPath & HourLog) For Output As 1
                Print #1, "---------------------------------------------------------------------------------------------"; vbCrLf; _
                          "Mail Robot HOURS ACTIVITY LOG - DATE: "; Format(Now, "dd/mm/yyyy HH:nn:ss"); vbCrLf; _
                          "---------------------------------------------------------------------------------------------"; vbCrLf; _
                          "DATA DATE          HOURS            FINISHED"; vbCrLf; _
                          "---------          -------          --------------------------"
        Close #1
        
        'To hours history, appending a line that separates the hours of today, from tomorrow:
        Open (strAppPath & HourHist) For Append As 1
            Print #1, "---------------------------------------------------------------------------------------------"
        Close #1
    End If
    
    'Data to be converted should only be of the past days:
    If DateValue(MyNow) >= Date Then
        If InStr(1, txtDaysLog.Text, "CONVERSION") = 0 Then
            txtDaysLog.Text = txtDaysLog.Text & vbCrLf & "*** CONVERSION PROCESS HELD UNTILL TARGET DATE IS MET ***"
            txtDaysLog.SelStart = Len(txtDaysLog.Text)
            txtDaysLog.SelLength = 1
        End If
        Exit Sub
    Else
        txtDaysLog.Text = Replace(txtDaysLog.Text, vbCrLf & "*** CONVERSION PROCESS HELD UNTILL TARGET DATE IS MET ***", "")
        txtDaysLog.SelStart = Len(txtDaysLog.Text)
        txtDaysLog.SelLength = 1
    End If
    
    'Clearing the days log, if met the date to clear files:
    If CDate(LogClearDate) = Date Then
        Kill (strAppPath & DayLog)
        Kill (strAppPath & DayHist)
        Kill (strAppPath & HourLog)
        Kill (strAppPath & HourHist)
        
        LogClearDate = DateAdd("d", strLogLife, Date)
    End If
    
    
    If DateValue(MyNow) > EndDate Then
        'Hooray! Done processing the entire group of days.
        tmrChecker.Enabled = False
        cmdAction.Caption = "&START"
        
        'Re-enabling controls
        Me.Frame1.Enabled = True
        Me.Frame2.Enabled = True
        Me.Frame3.Enabled = True
        Me.chkFilterON.Enabled = True
        Me.txtFilter.Enabled = True
        Me.cmdFilter.Enabled = True
        Me.cmdEdit.Enabled = True
        Me.comHours.Enabled = True
    Else
        If boStopTimer Then           'This is True if user has chosen to STOP data process.
            tmrChecker.Enabled = False
            txtDaysLog.Text = txtDaysLog.Text & vbCrLf & "**CONVERSION PROCESS INTERRUPTED BY USER**"
        Else
            DayBegin = Now
            Open (strAppPath & DayLog) For Append As 1
                Print #1, Format(MyNow, "dd/mm/yy"); "       "; Format(DayBegin, "dd/mm/yy hh:nn"); "      IN PROGRESS         N/A"
            Close #1
            Me.txtDaysLog.Text = Me.txtDaysLog.Text & vbCrLf & Format(MyNow, "dd/mm/yy") & "       " & Format(DayBegin, "dd/mm/yy hh:nn") & "      IN PROGRESS         N/A"
            
            
            Call TheLister
            'Scroll down:
            txtDaysLog.SelStart = Len(txtDaysLog.Text)
            txtDaysLog.SelLength = 1
        End If
    End If
    
End Sub

Private Sub tmrFreeSpacer_Timer()
    On Error Resume Next
    Dim DelPath, DelYear, DelMonth, DelDay As String
    Dim FoundMonth, FoundDay As Boolean
    Dim Percentage As Integer
    Dim i As Integer
    
    Call GetDiskSpace(txtDest.Text)
    
    Percentage = dblFreeSpace / dblTotalSize * 100
    
    lblStatus.Caption = "Total disk size: " & dblTotalSize & " GB  --  Free space: " & dblFreeSpace & " GB  --  Percentage: " & Percentage & " %"
    Exit Sub
    'Delete old data only if free space is approaching critical level...
    If Percentage > CriticalFreeSpace Then Exit Sub
    
    'The following block is about finding the path of the oldest .pst files, to delete them:
    'Finding the earliest year:
    For i = 0 To 5
        If Dir(strDest & Year(Date) - i, vbDirectory) = "" Then
            DelYear = Year(Date) - i - 1 & "\"
            Exit For
        End If
    Next i
    
    'Finding the earliest month in the earliest year:
    For i = 1 To 12
        If Dir(strDest & DelYear & i, vbDirectory) <> "" Then
            DelMonth = i & "\"
            FoundMonth = True
            Exit For
        End If
    Next i
    
    'If Year folder is empty, delete it and exit sub:
    If Not FoundMonth Then
        RmDir (strDest & strYear)
        Exit Sub
    End If
    
    'Finding the earliest day in the earliest month:
    For i = 1 To 31
        If Dir(strDest & DelYear & DelMonth & i, vbDirectory) <> "" Then
            DelDay = i & "\"
            FoundDay = True
            Exit For
        End If
    Next i
    
    'If Month folder was empty, delete it
    If Not FoundDay Then
        RmDir (strDest & DelYear & DelMonth)
        Exit Sub
    End If
    
    'Constructing full path of the oldest data to be deleted:
    DelPath = strDest & DelYear & DelMonth & DelDay
    
    'Deleting old data and the folder that holds them:
    Kill DelPath & "*.*"
    RmDir strDest & DelYear & DelMonth & DelDay
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

