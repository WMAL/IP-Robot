Attribute VB_Name = "mdlDiskStats"
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

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As LARGE_INTEGER, lpTotalNumberOfBytes As LARGE_INTEGER, lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Long

Public Sub GetDiskSpace(ByVal sDrive As String)
    Dim lResult As Long
    Dim liAvailable As LARGE_INTEGER
    Dim liTotal As LARGE_INTEGER
    Dim liFree As LARGE_INTEGER
    Dim dblAvailable As Double
    Dim dblTotal As Double
    Dim dblFree As Double
    
    If Right(sDrive, 1) <> "" Then sDrive = sDrive & ""
    'Determine the Available Space, Total Size and Free Space of a drive
    lResult = GetDiskFreeSpaceEx(sDrive, liAvailable, liTotal, liFree)
   
    'Convert the return values from LARGE_INTEGER to doubles
    dblAvailable = CLargeInt(liAvailable.lowpart, liAvailable.highpart)
    dblTotal = CLargeInt(liTotal.lowpart, liTotal.highpart)
    dblFree = CLargeInt(liFree.lowpart, liFree.highpart)
    
    dblFreeSpace = Round((dblFree / 1000000000) * 0.93, 1)
    dblTotalSize = Round((dblTotal / 1000000000) * 0.93, 1)
    
End Sub

Private Function CLargeInt(Lo As Long, Hi As Long) As Double
    'This function converts the LARGE_INTEGER data type to a double
    Dim dblLo As Double, dblHi As Double
   
    If Lo < 0 Then
        dblLo = 2 ^ 32 + Lo
    Else
        dblLo = Lo
    End If
   
    If Hi < 0 Then
        dblHi = 2 ^ 32 + Hi
    Else
        dblHi = Hi
    End If
    CLargeInt = dblLo + dblHi * 2 ^ 32
End Function

