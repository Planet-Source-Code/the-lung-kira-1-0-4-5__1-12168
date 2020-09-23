Attribute VB_Name = "shlobj"
Option Explicit


Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidl As Long, ByVal pszPath As String) As Boolean
Public Declare Function SHGetSpecialFolderPath Lib "shell32" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Boolean) As Boolean


    Public BROWSEINFO As BROWSEINFO
    Public Type BROWSEINFO
        hwndOwner As Long
        pidlRoot As Long 'ITEMIDLIST
        pszDisplayName As String
        lpszTitle As String
        ulFlags As Long
        lpfn As Long
        lParam As Long
        iImage As Long
    End Type


    Public Const BIF_RETURNONLYFSDIRS = &H1         'For finding a folder to start document searching
    Public Const BIF_DONTGOBELOWDOMAIN = &H2        'For starting the Find Computer
    Public Const BIF_STATUSTEXT = &H4               'Top of the dialog has 2 lines of text for BROWSEINFO.lpszTitle and one line if this flag is set.  Passing the message BFFM_SETSTATUSTEXTA to the hwnd can set the rest of the text.  This is not used with BIF_USENEWUI and BROWSEINFO.lpszTitle gets all three lines of text.
    Public Const BIF_RETURNFSANCESTORS = &H8
    Public Const BIF_EDITBOX = &H10                 'Add an editbox to the dialog
    Public Const BIF_VALIDATE = &H20                'insist on valid result (or CANCEL)
    Public Const BIF_NEWDIALOGSTYLE = &H40          'Use the new dialog layout with the ability to resize Caller needs to call OleInitialize() before using this API
    Public Const BIF_USENEWUI = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)
    Public Const BIF_BROWSEINCLUDEURLS = &H80       'Allow URLs to be displayed or entered. (Requires BIF_USENEWUI)
    Public Const BIF_BROWSEFORCOMPUTER = &H1000     'Browsing for Computers.
    Public Const BIF_BROWSEFORPRINTER = &H2000      'Browsing for Printers
    Public Const BIF_BROWSEINCLUDEFILES = &H4000    'Browsing for Everything
    Public Const BIF_SHAREABLE = &H8000             'sharable resources displayed (remote shares, requires BIF_USENEWUI)

    Public Const CSIDL_DESKTOP = &H0                             '<desktop>
    Public Const CSIDL_INTERNET = &H1                            'Internet Explorer (icon on desktop)
    Public Const CSIDL_PROGRAMS = &H2                            'Start Menu\Programs
    Public Const CSIDL_CONTROLS = &H3                            'My Computer\Control Panel
    Public Const CSIDL_PRINTERS = &H4                            'My Computer\Printers
    Public Const CSIDL_PERSONAL = &H5                            'My Documents
    Public Const CSIDL_FAVORITES = &H6                           '<user name>\Favorites
    Public Const CSIDL_STARTUP = &H7                             'Start Menu\Programs\Startup
    Public Const CSIDL_RECENT = &H8                              '<user name>\Recent
    Public Const CSIDL_SENDTO = &H9                              '<user name>\SendTo
    Public Const CSIDL_BITBUCKET = &HA                           '<desktop>\Recycle Bin
    Public Const CSIDL_STARTMENU = &HB                           '<user name>\Start Menu
    Public Const CSIDL_DESKTOPDIRECTORY = &H10                   '<user name>\Desktop
    Public Const CSIDL_DRIVES = &H11                             'My Computer
    Public Const CSIDL_NETWORK = &H12                            'Network Neighborhood
    Public Const CSIDL_NETHOOD = &H13                            '<user name>\nethood
    Public Const CSIDL_FONTS = &H14                              'windows\fonts
    Public Const CSIDL_TEMPLATES = &H15
    Public Const CSIDL_COMMON_STARTMENU = &H16                   'All Users\Start Menu
    Public Const CSIDL_COMMON_PROGRAMS = &H17                     'All Users\Programs
    Public Const CSIDL_COMMON_STARTUP = &H18                     'All Users\Startup
    Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19            'All Users\Desktop
    Public Const CSIDL_APPDATA = &H1A                            '<user name>\Application Data
    Public Const CSIDL_PRINTHOOD = &H1B                          '<user name>\PrintHood
    Public Const CSIDL_LOCAL_APPDATA = &H1C                      '<user name>\Local Settings\Applicaiton Data (non roaming)
    Public Const CSIDL_ALTSTARTUP = &H1D                         'non localized startup
    Public Const CSIDL_COMMON_ALTSTARTUP = &H1E                  'non localized common startup
    Public Const CSIDL_COMMON_FAVORITES = &H1F
    Public Const CSIDL_INTERNET_CACHE = &H20
    Public Const CSIDL_COOKIES = &H21
    Public Const CSIDL_HISTORY = &H22
    Public Const CSIDL_COMMON_APPDATA = &H23                     'All Users\Application Data
    Public Const CSIDL_WINDOWS = &H24                            'GetWindowsDirectory()
    Public Const CSIDL_SYSTEM = &H25                             'GetSystemDirectory()
    Public Const CSIDL_PROGRAM_FILES = &H26                      'C:\Program Files
    Public Const CSIDL_MYPICTURES = &H27                         'C:\Program Files\My Pictures
    Public Const CSIDL_PROFILE = &H28                            'USERPROFILE
    Public Const CSIDL_SYSTEMX86 = &H29                          'x86 system directory on RISC
    Public Const CSIDL_PROGRAM_FILESX86 = &H2A                   'x86 C:\Program Files on RISC
    Public Const CSIDL_PROGRAM_FILES_COMMON = &H2B               'C:\Program Files\Common
    Public Const CSIDL_PROGRAM_FILES_COMMONX86 = &H2C            'x86 Program Files\Common on RISC
    Public Const CSIDL_COMMON_TEMPLATES = &H2D                   'All Users\Templates
    Public Const CSIDL_COMMON_DOCUMENTS = &H2E                   'All Users\Documents
    Public Const CSIDL_COMMON_ADMINTOOLS = &H2F                  'All Users\Start Menu\Programs\Administrative Tools
    Public Const CSIDL_ADMINTOOLS = &H30                         '<user name>\Start Menu\Programs\Administrative Tools
    Public Const CSIDL_CONNECTIONS = &H31                        'Network and Dial-up Connections
    
    Public Const CSIDL_FLAG_CREATE = &H8000                      'combine with CSIDL_ value to force folder creation in SHGetFolderPath()
    Public Const CSIDL_FLAG_DONT_VERIFY = &H4000                 'combine with CSIDL_ value to return an unverified folder path
    Public Const CSIDL_FLAG_MASK = &HFF00                        'mask for all possible flag values

Public Function GetDirectory(hwnd As Long, strWindowTitle As String, bufFileName As String)
    With BROWSEINFO
        .hwndOwner = hwnd
        .pszDisplayName = Space$(MAX_PATH) 'Nothin special
        .lpszTitle = strWindowTitle & Chr$(0)
        .ulFlags = BIF_EDITBOX
    End With
    
    Dim tmpLong As Long
    
    tmpLong = SHBrowseForFolder(BROWSEINFO)
    If tmpLong Then
        Dim tmpString As String
        tmpString = Space$(MAX_PATH) 'Max dir length
        
        If SHGetPathFromIDList(tmpLong, tmpString) = True Then
            bufFileName = Fix_NullTermStr(tmpString) 'Send out
        Else
            Failed "SHGetPathFromIDList"
        End If
    End If
End Function
