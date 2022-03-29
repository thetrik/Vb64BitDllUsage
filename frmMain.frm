VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "64bit dll usage by The trick"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtInfo 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   60
      Width           =   4635
   End
   Begin VB.TextBox txtPID 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Text            =   "0"
      Top             =   2040
      Width           =   4695
   End
   Begin VB.CommandButton cmdGetInfo 
      Caption         =   "Get info using 64 bit dll"
      Enabled         =   0   'False
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   2460
      Width           =   3615
   End
   Begin VB.Label lblPID 
      Caption         =   "PID:"
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   1740
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // 64bit dll working demonstartion (loading / calling function)
' // By The trick, 2022
' //

Option Explicit

Private Type PROCESS_MEMORY_COUNTERS64
    cb                          As Long
    PageFaultCount              As Long
    PeakWorkingSetSize          As Currency
    WorkingSetSize              As Currency
    QuotaPeakPagedPoolUsage     As Currency
    QuotaPagedPoolUsage         As Currency
    QuotaPeakNonPagedPoolUsage  As Currency
    QuotaNonPagedPoolUsage      As Currency
    PagefileUsage               As Currency
    PeakPagefileUsage           As Currency
End Type
                         
Private Type UNICODE_STRING64
    Length                          As Integer
    MaxLength                       As Integer
    lPad                            As Long
    lpBuffer                        As Currency
End Type

Private Type ANSI_STRING64
    Length                          As Integer
    MaxLength                       As Integer
    lPad                            As Long
    lpBuffer                        As Currency
End Type

Private Declare Function StrFormatKBSize Lib "shlwapi" _
                         Alias "StrFormatKBSizeW" ( _
                         ByVal qdw As Currency, _
                         ByVal pszBuf As Long, _
                         ByVal cchBuf As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60.dll" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long

Private m_h64Lib    As Currency
Private m_p64Fn     As Currency

Private Sub Form_Load()
    Dim h64NtDll        As Currency
    Dim p64LdrLoad      As Currency
    Dim p64LdrGetProc   As Currency
    Dim sDllPath        As String
    Dim sFnName         As String
    Dim tDllPath        As UNICODE_STRING64
    Dim tFnName         As ANSI_STRING64
    Dim lStatus         As Long
    
    On Error GoTo err_handler
    
    If Not modX64Call.Initialize() Then
        MsgBox "Unable to initialize x64 call module", vbCritical
        Exit Sub
    End If
    
    h64NtDll = GetModuleHandle64("ntdll.dll")
    If h64NtDll = 0 Then
        MsgBox "Unable to get ntdll64", vbCritical
        Exit Sub
    End If
    
    p64LdrLoad = GetProcAddress64(h64NtDll, "LdrLoadDll")
    If p64LdrLoad = 0 Then
        MsgBox "Unable to get LdrLoadDll", vbCritical
        Exit Sub
    End If
    
    p64LdrGetProc = GetProcAddress64(h64NtDll, "LdrGetProcedureAddress")
    If p64LdrGetProc = 0 Then
        MsgBox "Unable to get LdrGetProcedureAddress", vbCritical
        Exit Sub
    End If
    
    ' // Load 64 bit dll
    sDllPath = App.Path & "\dll\dll.dll"
    
    tDllPath.Length = LenB(sDllPath)
    tDllPath.MaxLength = tDllPath.Length + 2
    GetMem4 StrPtr(sDllPath), tDllPath.lpBuffer
    
    GetMem4 CallX64(p64LdrLoad, 0&, 0&, VarPtr(tDllPath), VarPtr(m_h64Lib)), lStatus
    
    If lStatus < 0 Then
        MsgBox "Unable to load dll " & lStatus, vbCritical
        Exit Sub
    End If
    
    ' // GetProcAddress
    sFnName = StrConv("GetProcessMemoryInfo64", vbFromUnicode)
    
    tFnName.Length = LenB(sFnName)
    tFnName.MaxLength = tFnName.Length + 1
    GetMem4 StrPtr(sFnName), tFnName.lpBuffer
    
    GetMem4 CallX64(p64LdrGetProc, m_h64Lib, VarPtr(tFnName), 0, VarPtr(m_p64Fn)), lStatus
    
    If lStatus < 0 Then
        MsgBox "Unable to get procedure address " & lStatus, vbCritical
        Exit Sub
    End If
    
    cmdGetInfo.Enabled = True
    
    Exit Sub
    
err_handler:
    
    MsgBox "An error occured " & Err.Number, vbCritical
    
End Sub

Private Sub cmdGetInfo_Click()
    Dim lPID        As Long
    Dim tMemInfo    As PROCESS_MEMORY_COUNTERS64
    
    On Error GoTo err_handler
    
    lPID = Val(txtPID.Text)
    
    ' // Call 64 bit function
    tMemInfo.cb = LenB(tMemInfo)
    
    If CallX64(m_p64Fn, lPID, VarPtr(tMemInfo)) = 0 Then
        MsgBox "GetProcessMemoryInfo64 failed", vbInformation
    Else
        With tMemInfo
            txtInfo.Text = "WorkingSetSize: " & FormatSize(.WorkingSetSize) & vbNewLine & _
                            "PagefileUsage: " & FormatSize(.PagefileUsage) & vbNewLine & _
                            "PageFaultCount: " & .PageFaultCount & vbNewLine & _
                            "PeakPagefileUsage: " & FormatSize(.PeakPagefileUsage) & vbNewLine & _
                            "PeakWorkingSetSize: " & FormatSize(.PeakWorkingSetSize) & vbNewLine & _
                            "QuotaNonPagedPoolUsage: " & FormatSize(.QuotaNonPagedPoolUsage) & vbNewLine & _
                            "QuotaPagedPoolUsage: " & FormatSize(.QuotaPagedPoolUsage) & vbNewLine & _
                            "QuotaPeakNonPagedPoolUsage: " & FormatSize(.QuotaPeakNonPagedPoolUsage) & vbNewLine & _
                            "QuotaPeakPagedPoolUsage: " & FormatSize(.QuotaPeakPagedPoolUsage)
        End With
    End If
    
    Exit Sub
    
err_handler:
    
    MsgBox "An error occured " & Err.Number, vbCritical
    
End Sub

Private Function FormatSize( _
                 ByVal cValue As Currency) As String
    
    FormatSize = Space$(32)
    
    If StrFormatKBSize(cValue, StrPtr(FormatSize), Len(FormatSize)) Then
        FormatSize = Left$(FormatSize, InStr(1, FormatSize, vbNullChar) - 1)
    Else
        FormatSize = "ERROR"
    End If
    
End Function

Private Sub Form_Unload( _
            ByRef Cancel As Integer)
    modX64Call.Uninitialize
End Sub
