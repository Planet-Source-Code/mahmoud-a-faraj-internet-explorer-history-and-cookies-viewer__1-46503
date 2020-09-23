Attribute VB_Name = "Module1"
Public Const ERROR_CACHE_FIND_FAIL As Long = 0
Public Const ERROR_CACHE_FIND_SUCCESS As Long = 1
Public Const ERROR_FILE_NOT_FOUND As Long = 2
Public Const ERROR_ACCESS_DENIED As Long = 5
Public Const ERROR_INSUFFICIENT_BUFFER As Long = 122
Public Const MAX_PATH As Long = 260
Public Const MAX_CACHE_ENTRY_INFO_SIZE As Long = 4096

Public Const LMEM_FIXED As Long = &H0
Public Const LMEM_ZEROINIT As Long = &H40
Public Const LPTR As Long = (LMEM_FIXED Or LMEM_ZEROINIT)

Public Const NORMAL_CACHE_ENTRY As Long = &H1
Public Const EDITED_CACHE_ENTRY As Long = &H8
Public Const TRACK_OFFLINE_CACHE_ENTRY As Long = &H10
Public Const TRACK_ONLINE_CACHE_ENTRY As Long = &H20
Public Const STICKY_CACHE_ENTRY As Long = &H40
Public Const SPARSE_CACHE_ENTRY As Long = &H10000
Public Const COOKIE_CACHE_ENTRY As Long = &H100000
Public Const URLHISTORY_CACHE_ENTRY As Long = &H200000
Public Const URLCACHE_FIND_DEFAULT_FILTER As Long = NORMAL_CACHE_ENTRY Or _
                                                    COOKIE_CACHE_ENTRY Or _
                                                    URLHISTORY_CACHE_ENTRY Or _
                                                    TRACK_OFFLINE_CACHE_ENTRY Or _
                                                    TRACK_ONLINE_CACHE_ENTRY Or _
                                                    STICKY_CACHE_ENTRY
Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type
Private Type FILETIME
     dwLowDateTime As Long
     dwHighDateTime As Long
End Type

Private Type INTERNET_CACHE_ENTRY_INFO
     dwStructSize As Long
     lpszSourceUrlName As Long
     lpszLocalFileName As Long
     CacheEntryType As Long
     dwUseCount As Long
     dwHitRate As Long
     dwSizeLow As Long
     dwSizeHigh As Long
     LastModifiedTime As FILETIME
     ExpireTime As FILETIME
     LastAccessTime As FILETIME
     LastSyncTime As FILETIME
     lpHeaderInfo As Long
     dwHeaderInfoSize As Long
     lpszFileExtension As Long
     dwExemptDelta  As Long
End Type
Public Type Internet_Cache_Entry
     'dwStructSize As Long
     SourceUrlName As String
     LocalFileName As String
     'CacheEntryType  As Long
     UseCount As Long
     HitRate As Long
     Size As Long
     'dwSizeHigh As Long
     LastModifiedTime As Date
     ExpireTime As Date
     LastAccessTime As Date
     LastSyncTime As Date
     HeaderInfo As String
     'dwHeaderInfoSize As Long
     FileExtension As String
     'ExemptDelta  As Long
End Type

'==============================================================================
'   Déclarations API

Private Declare Function FileTimeToLocalFileTime Lib "KERNEL32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "KERNEL32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "KERNEL32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "KERNEL32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long

Private Declare Function FindFirstUrlCacheEntry Lib "Wininet.dll" _
     Alias "FindFirstUrlCacheEntryA" _
    (ByVal lpszUrlSearchPattern As String, _
     lpFirstCacheEntryInfo As Any, _
     lpdwFirstCacheEntryInfoBufferSize As Long) As Long

Private Declare Function FindNextUrlCacheEntry Lib "Wininet.dll" _
     Alias "FindNextUrlCacheEntryA" _
    (ByVal hEnumHandle As Long, _
     lpNextCacheEntryInfo As Any, _
     lpdwNextCacheEntryInfoBufferSize As Long) As Long

Private Declare Function FindCloseUrlCache Lib "Wininet.dll" _
     (ByVal hEnumHandle As Long) As Long

Public Declare Function DeleteUrlCacheEntry Lib "Wininet.dll" _
     Alias "DeleteUrlCacheEntryA" _
    (ByVal lpszUrlName As String) As Long
     
Private Declare Sub CopyMemory Lib "KERNEL32" _
     Alias "RtlMoveMemory" _
     (pDest As Any, _
    pSource As Any, _
    ByVal dwLength As Long)

Private Declare Function lstrcpyA Lib "KERNEL32" _
    (ByVal RetVal As String, ByVal Ptr As Long) As Long
                        
Private Declare Function lstrlenA Lib "KERNEL32" _
    (ByVal Ptr As Any) As Long
    
Private Declare Function LocalAlloc Lib "KERNEL32" _
     (ByVal uFlags As Long, _
    ByVal uBytes As Long) As Long
    
Private Declare Function LocalFree Lib "KERNEL32" _
     (ByVal hMem As Long) As Long
Public Function GetURLCache(URL() As Internet_Cache_Entry, URLHistory() As Internet_Cache_Entry, Cookies() As Internet_Cache_Entry)
     Dim ICEI As INTERNET_CACHE_ENTRY_INFO
     Dim hFile As Long
     Dim cachefile As String
     Dim posUrl As Long
     Dim posEnd As Long
     Dim dwBuffer As Long
     Dim pntrICE As Long
     
     dwBuffer = 0
     ReDim URL(0)
     ReDim URLHistory(0)
     ReDim Cookies(0)
     hFile = FindFirstUrlCacheEntry(0&, ByVal 0, dwBuffer)
     If (hFile = ERROR_CACHE_FIND_FAIL) And _
        (Err.LastDllError = ERROR_INSUFFICIENT_BUFFER) Then
        pntrICE = LocalAlloc(LMEM_FIXED, dwBuffer)
        If pntrICE Then
         CopyMemory ByVal pntrICE, dwBuffer, 4
         hFile = FindFirstUrlCacheEntry(vbNullString, ByVal pntrICE, dwBuffer)
         If hFile <> ERROR_CACHE_FIND_FAIL Then
            Do
                 CopyMemory ICEI, ByVal pntrICE, Len(ICEI)
                 If (ICEI.CacheEntryType And _
                     NORMAL_CACHE_ENTRY) = NORMAL_CACHE_ENTRY Then
                 Select Case ICEI.CacheEntryType
                    Case URLHISTORY_CACHE_ENTRY + NORMAL_CACHE_ENTRY
                    ReDim Preserve URLHistory(UBound(URLHistory) + 1)
                    URLHistory(UBound(URLHistory) - 1).SourceUrlName = GetStrFromPtrA(ICEI.lpszSourceUrlName)
                    URLHistory(UBound(URLHistory) - 1).LocalFileName = GetStrFromPtrA(ICEI.lpszLocalFileName)
                    URLHistory(UBound(URLHistory) - 1).FileExtension = GetStrFromPtrA(ICEI.lpszFileExtension)
                    URLHistory(UBound(URLHistory) - 1).HeaderInfo = GetStrFromPtrA(ICEI.lpHeaderInfo)
                    URLHistory(UBound(URLHistory) - 1).HitRate = ICEI.dwHitRate
                    URLHistory(UBound(URLHistory) - 1).ExpireTime = FileTime2SystemTime(ICEI.ExpireTime)
                    URLHistory(UBound(URLHistory) - 1).LastAccessTime = FileTime2SystemTime(ICEI.LastAccessTime)
                    URLHistory(UBound(URLHistory) - 1).LastModifiedTime = FileTime2SystemTime(ICEI.LastModifiedTime)
                    URLHistory(UBound(URLHistory) - 1).LastSyncTime = FileTime2SystemTime(ICEI.LastSyncTime)
                    URLHistory(UBound(URLHistory) - 1).Size = ICEI.dwSizeHigh * 2 ^ 32 + ICEI.dwSizeLow
                    URLHistory(UBound(URLHistory) - 1).UseCount = ICEI.dwUseCount
                    Case COOKIE_CACHE_ENTRY + NORMAL_CACHE_ENTRY
                    ReDim Preserve Cookies(UBound(Cookies) + 1)
                    Cookies(UBound(Cookies) - 1).SourceUrlName = GetStrFromPtrA(ICEI.lpszSourceUrlName)
                    Cookies(UBound(Cookies) - 1).LocalFileName = GetStrFromPtrA(ICEI.lpszLocalFileName)
                    Cookies(UBound(Cookies) - 1).FileExtension = GetStrFromPtrA(ICEI.lpszFileExtension)
                    Cookies(UBound(Cookies) - 1).HeaderInfo = GetStrFromPtrA(ICEI.lpHeaderInfo)
                    Cookies(UBound(Cookies) - 1).HitRate = ICEI.dwHitRate
                    Cookies(UBound(Cookies) - 1).ExpireTime = FileTime2SystemTime(ICEI.ExpireTime)
                    Cookies(UBound(Cookies) - 1).LastAccessTime = FileTime2SystemTime(ICEI.LastAccessTime)
                    Cookies(UBound(Cookies) - 1).LastModifiedTime = FileTime2SystemTime(ICEI.LastModifiedTime)
                    Cookies(UBound(Cookies) - 1).LastSyncTime = FileTime2SystemTime(ICEI.LastSyncTime)
                    Cookies(UBound(Cookies) - 1).Size = ICEI.dwSizeHigh * 2 ^ 32 + ICEI.dwSizeLow
                    Cookies(UBound(Cookies) - 1).UseCount = ICEI.dwUseCount
                    Case Else
                    ReDim Preserve URL(UBound(URL) + 1)
                    URL(UBound(URL) - 1).SourceUrlName = GetStrFromPtrA(ICEI.lpszSourceUrlName)
                    URL(UBound(URL) - 1).LocalFileName = GetStrFromPtrA(ICEI.lpszLocalFileName)
                    URL(UBound(URL) - 1).FileExtension = GetStrFromPtrA(ICEI.lpszFileExtension)
                    URL(UBound(URL) - 1).HeaderInfo = GetStrFromPtrA(ICEI.lpHeaderInfo)
                    URL(UBound(URL) - 1).HitRate = ICEI.dwHitRate
                    URL(UBound(URL) - 1).ExpireTime = FileTime2SystemTime(ICEI.ExpireTime)
                    URL(UBound(URL) - 1).LastAccessTime = FileTime2SystemTime(ICEI.LastAccessTime)
                    URL(UBound(URL) - 1).LastModifiedTime = FileTime2SystemTime(ICEI.LastModifiedTime)
                    URL(UBound(URL) - 1).LastSyncTime = FileTime2SystemTime(ICEI.LastSyncTime)
                    URL(UBound(URL) - 1).Size = ICEI.dwSizeHigh * 2 ^ 32 + ICEI.dwSizeLow
                    URL(UBound(URL) - 1).UseCount = ICEI.dwUseCount
               
                 End Select
                 End If
                 Call LocalFree(pntrICE)
                 dwBuffer = 0
                 Call FindNextUrlCacheEntry(hFile, ByVal 0, dwBuffer)
                 pntrICE = LocalAlloc(LMEM_FIXED, dwBuffer)
                 CopyMemory ByVal pntrICE, dwBuffer, 4
            Loop While FindNextUrlCacheEntry(hFile, ByVal pntrICE, dwBuffer)
         End If 'hFile
        End If 'pntrICE
     End If 'hFile
     Call LocalFree(pntrICE)
     Call FindCloseUrlCache(hFile)
End Function

Private Function GetStrFromPtrA(ByVal lpszA As Long) As String
     GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
     Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
End Function

Private Function FileTime2SystemTime(FileT As FILETIME) As Date
Dim SysT As SYSTEMTIME
FileTimeToLocalFileTime FileT, FileT
FileTimeToSystemTime FileT, SysT
FileTime2SystemTime = TimeSerial(SysT.wHour, SysT.wMinute, SysT.wSecond) + DateSerial(SysT.wYear, SysT.wMonth, SysT.wDay)
End Function

Public Function DeleteUrlCache(liste() As Internet_Cache_Entry) As Boolean
Dim x As Long

For x = LBound(liste) To UBound(liste) - 1
DeleteUrlCache = DeleteUrlCacheEntry(liste(x).SourceUrlName)
Next x
End Function



Public Function deleteselecteditem(selecteditem$) As Boolean

deleteselecteditem = DeleteUrlCacheEntry(selecteditem)
 
End Function
