VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Patcher"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   60
      TabIndex        =   0
      Top             =   1290
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "&No Handler"
      TabPicture(0)   =   "frmMain.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstFileType(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCreate"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Text Handler"
      TabPicture(1)   =   "frmMain.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstFileType(1)"
      Tab(1).Control(1)=   "cmdRemove(0)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Custom Handler"
      TabPicture(2)   =   "frmMain.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstFileType(2)"
      Tab(2).Control(1)=   "cmdRemove(1)"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   345
         Index           =   1
         Left            =   -71250
         TabIndex        =   6
         Top             =   480
         Width           =   1245
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   345
         Index           =   0
         Left            =   -71250
         TabIndex        =   5
         Top             =   480
         Width           =   1245
      End
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Create"
         Height          =   345
         Left            =   3750
         TabIndex        =   4
         Top             =   480
         Width           =   1245
      End
      Begin VB.ListBox lstFileType 
         Height          =   4560
         Index           =   2
         Left            =   -74910
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   390
         Width           =   3375
      End
      Begin VB.ListBox lstFileType 
         Height          =   4560
         Index           =   1
         Left            =   -74910
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   390
         Width           =   3375
      End
      Begin VB.ListBox lstFileType 
         Height          =   4560
         Index           =   0
         Left            =   90
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   390
         Width           =   3375
      End
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmMain.frx":0496
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   90
      TabIndex        =   7
      Top             =   30
      Width           =   5235
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Created 06/01/05 by Richard Mewett
'This program is used to associate filetypes with the Text Filter to enable the Windows
'Search Tool to scan them.

'For example .BAS, .FRM & .VBP files could be searched on Windows 98 but on Windows XP
'they will be ignored even when using *.* as a mask.

'########################################################################################################
'Windows API Declarations
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Const HKEY_CLASSES_ROOT = &H80000000

Private Const REG_SZ = 1
Private Const REG_BINARY = 3

Private Const REG_OPTION_NON_VOLATILE = 0
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SYNCHRONIZE = &H100000
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Const ERROR_NO_MORE_ITEMS As Long = 259
Private Const BUFFER_SIZE As Long = 255

'########################################################################################################
'Applications Declarations
Private Const TEXT_HANDLER_CLSID = "{5e941d80-bf96-11cd-b579-08002b30bfeb}"

Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    'retrieve nformation about the key
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            'Create a buffer
            strBuf = String(lDataBufSize, Chr$(0))
            'retrieve the key's content
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                'Remove the unnecessary chr$(0)'s
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim strData As Integer
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
        End If
    End If
End Function
Private Sub LoadFileTypes()
    Dim hKey As Long
    Dim hFTKey As Long
    Dim lBufferSize As Long
    Dim nCount As Integer
    Dim sName As String
    
    '#################################################################################
    'This code loads a list of all filetypes that are registered on the system.
    
    'The program used 3 lists to load the filetypes:
    'List 0 - No Filter
    'List 1 - Associated with the Text Filter
    'List 2 - Associated with a Custom Filter (i.e. Imagess, Word Processor Documents)
    '#################################################################################
    
    lBufferSize = BUFFER_SIZE
    
    'Open the root registry key
    If RegOpenKey(HKEY_CLASSES_ROOT, "", hKey) = 0 Then
        'Enumerate the keys to load every FileType
        sName = Space(BUFFER_SIZE)
        While RegEnumKeyEx(hKey, nCount, sName, lBufferSize, ByVal 0&, vbNullString, ByVal 0&, ByVal 0&) <> ERROR_NO_MORE_ITEMS
            If Mid$(sName, 1, 1) = "." Then
                'It is a filetype since it starts with a "." character
                If RegOpenKey(HKEY_CLASSES_ROOT, sName, hFTKey) = 0 Then
                    'Check the current Filter
                    If RegOpenKey(hFTKey, "PersistentHandler", hFTKey) = 0 Then
                        If RegQueryStringValue(hFTKey, "") = TEXT_HANDLER_CLSID Then
                            lstFileType(1).AddItem sName
                        Else
                            lstFileType(2).AddItem sName
                        End If
                    Else
                        lstFileType(0).AddItem sName
                    End If
                End If
                
                'Close the key for the filetype
                RegCloseKey hFTKey
            End If
            
            nCount = nCount + 1
            sName = Space$(BUFFER_SIZE)
            lBufferSize = BUFFER_SIZE
        Wend
        
        'Close the root key
        RegCloseKey hKey
        
        'Set each list to select the first entry (if entries exist)
        For nCount = lstFileType.LBound To lstFileType.uBound
            If lstFileType(nCount).ListCount > 0 Then
                lstFileType(nCount).ListIndex = 0
            End If
        Next nCount
    End If
End Sub

Private Sub cmdCreate_Click()
    Dim hKey As Long
    Dim nCount As Integer
    
    '#################################################################################
    'This code creates the registry entries for the selected filetypes to associate
    'them with the Text Filter
    '#################################################################################
    
    Screen.MousePointer = vbHourglass
    
    With lstFileType(0)
        For nCount = 0 To .ListCount - 1
            If .Selected(nCount) Then
                RegCreateKeyEx HKEY_CLASSES_ROOT, .List(nCount), 0, "REG_DWORD", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, 0&
                If hKey <> 0 Then
                    RegCreateKeyEx hKey, "PersistentHandler", 0, "REG_DWORD", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ByVal 0&, hKey, 0&
                End If
                If hKey <> 0 Then
                    RegSetValueEx hKey, "", 0, REG_SZ, ByVal TEXT_HANDLER_CLSID, Len(TEXT_HANDLER_CLSID)
                End If
                RegCloseKey hKey
            End If
        Next nCount
    End With
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdRemove_Click(Index As Integer)
    Dim hKey As Long
    Dim nCount As Integer
    
    '#################################################################################
    'This code removes the registry entries for the selected filetypes which associate
    'them with a Filter
    '#################################################################################
    
    Screen.MousePointer = vbHourglass
    
    With lstFileType(Index + 1)
        For nCount = 0 To .ListCount - 1
            If .Selected(nCount) Then
                RegOpenKey HKEY_CLASSES_ROOT, .List(nCount), hKey
                If hKey <> 0 Then
                    RegDeleteKey hKey, "PersistentHandler"
                End If
                RegCloseKey hKey
            End If
        Next nCount
    End With
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
    LoadFileTypes
End Sub


