Attribute VB_Name = "modFolderBrowse"
Option Explicit
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Const ERROR_SUCCESS As Long = 0
Private Const MAX_PATH As Long = 260
Private Const CSIDL_NETWORK As Long = &H12
Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const BIF_BROWSEFORCOMPUTER As Long = &H1000

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Public Declare Function SHBrowseForFolder Lib "shell32.dll" (lpbi As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Function GetBrowseNetworkWorkstation(Owner As Long) As String
  'This Function from www.mvps.org/vbnet
  'returns only a valid network server or
  'workstation (does not display the shares)
   Dim BI As BROWSEINFO
   Dim pidl As Long
   Dim sPath As String
   Dim pos As Integer
   
   
  'obtain the pidl to the special folder 'network'
   If SHGetSpecialFolderLocation(Owner, _
                                 CSIDL_NETWORK, _
                                 pidl) = ERROR_SUCCESS Then
       
     'fill in the required members, limiting the
     'Browse to the network by specifying the
     'returned pidl as pidlRoot
      With BI
         .hOwner = Owner
         .pidlRoot = pidl
         .pszDisplayName = Space$(MAX_PATH)
         .lpszTitle = "Select a network computer."
         .ulFlags = BIF_BROWSEFORCOMPUTER
      End With
         
     'show the browse dialog. We don't need
     'a pidl, so it can be used in the If..then directly.
      If SHBrowseForFolder(BI) <> 0 Then
               
         'a server was selected. Although a valid pidl
         'is returned, SHGetPathFromIDList only return
         'paths to valid file system objects, of which
         'a networked machine is not. However, the
         'BROWSEINFO displayname member does contain
         'the selected item, which we return
          GetBrowseNetworkWorkstation = TrimNull(BI.pszDisplayName)
            
      End If  'If SHBrowseForFolder
      
      Call CoTaskMemFree(pidl)
               
   End If  'If SHGetSpecialFolderLocation
   
End Function
'****************************************************************************
' Trim to first Null character
' Inputs: String with null characaters
' Return: String up to where first null character occured
'****************************************************************************
Public Function TrimNull(item As String) As String
    Dim pos As Integer
        pos = InStr(item, Chr$(0))
        If pos Then item = Left$(item, pos - 1)
        TrimNull = item
End Function

