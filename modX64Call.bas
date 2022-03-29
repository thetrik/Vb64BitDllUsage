Attribute VB_Name = "modX64Call"

' //
' // modX64Call.bas
' // Module for calling functions in long-mode (x64)
' // by The trick 2018 - 2020
' //

Option Explicit

Private Const ProcessBasicInformation As Long = 0
Private Const MEM_RESERVE             As Long = &H2000&
Private Const MEM_COMMIT              As Long = &H1000&
Private Const MEM_RELEASE             As Long = &H8000&
Private Const PAGE_READWRITE          As Long = 4&
Private Const FADF_AUTO               As Long = 1
Private Const PAGE_EXECUTE_READWRITE  As Long = &H40&
Private Const PROCESS_VM_READ         As Long = &H10

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

Private Type PROCESS_BASIC_INFORMATION64
    ExitStatus                      As Long
    Reserved0                       As Long
    PebBaseAddress                  As Currency
    AffinityMask                    As Currency
    BasePriority                    As Long
    Reserved1                       As Long
    uUniqueProcessId                As Currency
    uInheritedFromUniqueProcessId   As Currency
End Type

Private Type IMAGE_DATA_DIRECTORY
    VirtualAddress                  As Long
    Size                            As Long
End Type

Private Type IMAGE_EXPORT_DIRECTORY
    Characteristics                 As Long
    TimeDateStamp                   As Long
    MajorVersion                    As Integer
    MinorVersion                    As Integer
    pName                           As Long
    Base                            As Long
    NumberOfFunctions               As Long
    NumberOfNames                   As Long
    AddressOfFunctions              As Long
    AddressOfNames                  As Long
    AddressOfNameOrdinals           As Long
End Type

Private Type SAFEARRAYBOUND
    cElements                       As Long
    lLbound                         As Long
End Type

Private Type SAFEARRAY1D
    cDims                           As Integer
    fFeatures                       As Integer
    cbElements                      As Long
    cLocks                          As Long
    pvData                          As Long
    Bounds                          As SAFEARRAYBOUND
End Type

Private Declare Function OpenProcess Lib "kernel32" ( _
                         ByVal dwDesiredAccess As Long, _
                         ByVal bInheritHandle As Long, _
                         ByVal dwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function CloseHandle Lib "kernel32" ( _
                         ByVal hObject As Long) As Long
Private Declare Function NtWow64QueryInformationProcess64 Lib "ntdll" ( _
                         ByVal hProcess As Long, _
                         ByVal ProcessInformationClass As Long, _
                         ByRef pProcessInformation As Any, _
                         ByVal uProcessInformationLength As Long, _
                         ByRef puReturnLength As Long) As Long
Private Declare Function NtWow64ReadVirtualMemory64 Lib "ntdll" ( _
                         ByVal hProcess As Long, _
                         ByVal p64Address As Currency, _
                         ByRef Buffer As Any, _
                         ByVal l64BufferLen As Currency, _
                         ByRef pl64ReturnLength As Currency) As Long
Private Declare Function GetMem8 Lib "msvbvm60" ( _
                         ByRef Src As Any, _
                         ByRef Dst As Any) As Long
Private Declare Function GetMem4 Lib "msvbvm60" ( _
                         ByRef Src As Any, _
                         ByRef Dst As Any) As Long
Private Declare Function GetMem2 Lib "msvbvm60" ( _
                         ByRef Src As Any, _
                         ByRef Dst As Any) As Long
Private Declare Function GetMem1 Lib "msvbvm60" ( _
                         ByRef Src As Any, _
                         ByRef Dst As Any) As Long
Private Declare Function VirtualAlloc Lib "kernel32" ( _
                         ByVal lpAddress As Long, _
                         ByVal dwSize As Long, _
                         ByVal flAllocationType As Long, _
                         ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" ( _
                         ByVal lpAddress As Long, _
                         ByVal dwSize As Long, _
                         ByVal dwFreeType As Long) As Long
Private Declare Function DispCallFunc Lib "oleaut32.dll" ( _
                         ByRef pvInstance As Any, _
                         ByVal oVft As Long, _
                         ByVal cc As Long, _
                         ByVal vtReturn As VbVarType, _
                         ByVal cActuals As Long, _
                         ByRef prgvt As Any, _
                         ByRef prgpvarg As Any, _
                         ByRef pvargResult As Variant) As Long
Private Declare Function lstrcmp Lib "kernel32" _
                         Alias "lstrcmpA" ( _
                         ByRef lpString1 As Any, _
                         ByRef lpString2 As Any) As Long
Private Declare Function lstrcmpi Lib "kernel32" _
                         Alias "lstrcmpiA" ( _
                         ByRef lpString1 As Any, _
                         ByRef lpString2 As Any) As Long
Private Declare Function ArrPtr Lib "msvbvm60" _
                         Alias "VarPtr" ( _
                         ByRef psa() As Any) As Long
Private Declare Sub MoveArray Lib "msvbvm60" _
                    Alias "__vbaAryMove" ( _
                    ByRef Destination() As Any, _
                    ByRef Source As Any)
                         
Private m_pCodeBuffer   As Long
Private m_hCurHandle    As Long

' // Initialize module
Public Function Initialize() As Boolean
    
    If m_pCodeBuffer = 0 Then
        
        m_hCurHandle = OpenProcess(PROCESS_VM_READ, 0, GetCurrentProcessId())
        
        If m_hCurHandle = 0 Then
            Exit Function
        End If
        
        ' // Temporary buffer for caller
        ' // Be careful it doesn't support threading
        ' // To support threading you should ensure atomic access to that buffer
        m_pCodeBuffer = VirtualAlloc(0, 4096, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
        
        If m_pCodeBuffer = 0 Then
            CloseHandle m_hCurHandle
            Exit Function
        End If
        
    End If
    
    Initialize = True
    
End Function

' // Uninitialize module
Public Sub Uninitialize()
    
    If m_hCurHandle Then
        CloseHandle m_hCurHandle
    End If
    
    If m_pCodeBuffer Then
        VirtualFree m_pCodeBuffer, 0, MEM_RELEASE
    End If
    
End Sub

' //
' // Call 64 bit function by pointer
' //
Public Function CallX64( _
                ByVal pfn64 As Currency, _
                ParamArray vArgs() As Variant) As Currency
    Dim bCode()     As Byte             ' // Array to map code
    Dim vArg        As Variant
    Dim lIndex      As Long
    Dim lByteIdx    As Long
    Dim lArgs       As Long
    Dim tArrDesc    As SAFEARRAY1D
    Dim vRet        As Variant
    Dim hr          As Long
    
    If m_pCodeBuffer = 0 Then
        
        ' // Isn't initialized
        Err.Raise 5
        Exit Function
        
    End If
    
    ' // Map array
    tArrDesc.cbElements = 1
    tArrDesc.cDims = 1
    tArrDesc.fFeatures = FADF_AUTO
    tArrDesc.Bounds.cElements = 4096
    tArrDesc.pvData = m_pCodeBuffer
    
    MoveArray bCode(), VarPtr(tArrDesc)
    
    ' // Make x64call
    
    ' // JMP FAR 33:ADDR
    bCode(0) = &HEA
    
    GetMem4 m_pCodeBuffer + 7, bCode(1)
    GetMem2 &H33, bCode(5)
    
    lByteIdx = 7
    
    ' // stack alignment
    
    ' // PUSH RBX
    ' // MOV RBX, SS
    ' // PUSH RBP
    ' // MOV RBP, RSP
    ' // AND ESP, 0xFFFFFFF0
    ' // SUB RSP, 0x28 + Args

    If UBound(vArgs) <= 3 Then
        lArgs = 4
    Else
        lArgs = ((UBound(vArgs) - 3) + 1) And &HFFFFFFFE
    End If
    
    lArgs = lArgs * 8 + &H20
    
    GetMem8 -140194732553717.1373@, bCode(lByteIdx):    lByteIdx = lByteIdx + 8
    GetMem8 26004001868.3011@, bCode(lByteIdx):         lByteIdx = lByteIdx + 6
    GetMem4 lArgs, bCode(lByteIdx):                     lByteIdx = lByteIdx + 4
    
    For Each vArg In vArgs
        
        Select Case VarType(vArg)
        Case vbLong, vbString, vbInteger, vbByte, vbBoolean

            Select Case lIndex
            Case 0: GetMem4 &HC1C748, bCode(lByteIdx):  lByteIdx = lByteIdx + 3
            Case 1: GetMem4 &HC2C748, bCode(lByteIdx):  lByteIdx = lByteIdx + 3
            Case 2: GetMem4 &HC0C749, bCode(lByteIdx):  lByteIdx = lByteIdx + 3
            Case 3: GetMem4 &HC1C749, bCode(lByteIdx):  lByteIdx = lByteIdx + 3
            Case Else
            
                GetMem4 &H2444C748, bCode(lByteIdx):    lByteIdx = lByteIdx + 4
                GetMem1 (lIndex - 4) * 8 + &H20, bCode(lByteIdx):   lByteIdx = lByteIdx + 1

            End Select
            
            Select Case VarType(vArg)
            Case vbLong, vbInteger, vbByte, vbBoolean
                GetMem4 CLng(vArg), bCode(lByteIdx):            lByteIdx = lByteIdx + 4
            Case vbString
                GetMem4 ByVal StrPtr(vArg), bCode(lByteIdx):    lByteIdx = lByteIdx + 4
            End Select
            
        Case vbCurrency
        
            Select Case lIndex
            Case 0: GetMem2 &HB948, bCode(lByteIdx):  lByteIdx = lByteIdx + 2
            Case 1: GetMem2 &HBA48, bCode(lByteIdx):  lByteIdx = lByteIdx + 2
            Case 2: GetMem2 &HB849, bCode(lByteIdx):  lByteIdx = lByteIdx + 2
            Case 3: GetMem2 &HB949, bCode(lByteIdx):  lByteIdx = lByteIdx + 2
            Case Else
            
                GetMem2 &HB848, bCode(lByteIdx):      lByteIdx = lByteIdx + 2
                GetMem8 CCur(vArg), bCode(lByteIdx):  lByteIdx = lByteIdx + 8
                GetMem4 &H24448948, bCode(lByteIdx):  lByteIdx = lByteIdx + 4
                GetMem1 (lIndex - 4) * 8 + &H20, bCode(lByteIdx):   lByteIdx = lByteIdx + 1
                
            End Select
            
            If lIndex < 4 Then
                GetMem8 CCur(vArg), bCode(lByteIdx):  lByteIdx = lByteIdx + 8
            End If
        
        Case Else
            
            Err.Raise 13
            Exit Function
            
        End Select
        
        lIndex = lIndex + 1
        
    Next
    
    ' // MOV RAX, pfn: CALL RAX
    GetMem2 &HB848, bCode(lByteIdx):    lByteIdx = lByteIdx + 2
    GetMem8 pfn64, bCode(lByteIdx):     lByteIdx = lByteIdx + 8
    GetMem2 &HD0FF&, bCode(lByteIdx):   lByteIdx = lByteIdx + 2
    
    ' // LEAVE
    ' // MOV SS, RBX
    ' // POP RBX
    GetMem8 39439134.1257@, bCode(lByteIdx):  lByteIdx = lByteIdx + 5
    
    ' // RAX to EAX/EDX pair
    ' // MOV RDX, RAX
    ' // SHR RDX, 0x20
    GetMem8 926531512503.7384@, bCode(lByteIdx):
    lByteIdx = lByteIdx + 7
    
    ' // JMP FAR 23:
    GetMem2 &H2DFF, bCode(lByteIdx):    lByteIdx = lByteIdx + 2
    GetMem4 0&, bCode(lByteIdx):        lByteIdx = lByteIdx + 4
    GetMem4 m_pCodeBuffer + lByteIdx + 6, bCode(lByteIdx)

    lByteIdx = lByteIdx + 4
    GetMem2 &H23&, bCode(lByteIdx):     lByteIdx = lByteIdx + 2

    bCode(lByteIdx) = &HC3

    hr = DispCallFunc(ByVal 0&, m_pCodeBuffer, 4, vbCurrency, 0, ByVal 0&, ByVal 0&, vRet)

    GetMem4 0&, ByVal ArrPtr(bCode)

    If hr < 0 Then
        Err.Raise hr
        Exit Function
    End If
    
    CallX64 = vRet
    
End Function

' //
' // Get procedure arrdess from 64 bit dll
' //
Public Function GetProcAddress64( _
                ByVal h64Lib As Currency, _
                ByRef sFunctionName As String) As Currency
    Dim lRvaNtHeaders       As Long
    Dim tExportData         As IMAGE_DATA_DIRECTORY
    Dim tExportDirectory    As IMAGE_EXPORT_DIRECTORY
    Dim lIndex              As Long
    Dim p64SymName          As Currency
    Dim tasFunction         As ANSI_STRING64
    Dim tasSymbol           As ANSI_STRING64
    Dim sAnsiString         As String
    Dim lOrdinal            As Long
    Dim p64Address          As Currency
    
    If h64Lib = 0 Then
        
        h64Lib = GetModuleHandle64(vbNullString)
            
        If h64Lib = 0 Then
        
            Err.Raise 5
            Exit Function
            
        End If
            
    End If

    sAnsiString = StrConv(sFunctionName, vbFromUnicode)
    
    GetMem4 StrPtr(sAnsiString), tasFunction.lpBuffer
    tasFunction.Length = LenB(sAnsiString)
    tasFunction.MaxLength = tasFunction.Length + 1
    
    ReadMem64 VarPtr(lRvaNtHeaders), h64Lib + 0.006@, Len(lRvaNtHeaders)
    ReadMem64 VarPtr(tExportData), h64Lib + lRvaNtHeaders / 10000 + 0.0136@, Len(tExportData)
    
    If tExportData.VirtualAddress = 0 Or tExportData.Size = 0 Then
        Err.Raise 453
        Exit Function
    End If
    
    ReadMem64 VarPtr(tExportDirectory), h64Lib + tExportData.VirtualAddress / 10000, Len(tExportDirectory)
    
    For lIndex = 0 To tExportDirectory.NumberOfNames - 1
        
        p64SymName = 0
        
        ReadMem64 VarPtr(p64SymName), (tExportDirectory.AddressOfNames + lIndex * 4) / 10000 + h64Lib, 4
        
        p64SymName = p64SymName + h64Lib
        
        tasSymbol.Length = StringLen64(p64SymName) * 10000
        tasSymbol.MaxLength = tasSymbol.Length
        tasSymbol.lpBuffer = p64SymName
        
        If CompareAnsiStrings64(tasFunction, tasSymbol, True) = 0 Then
            
            ReadMem64 VarPtr(lOrdinal), (tExportDirectory.AddressOfNameOrdinals + lIndex * 2) / 10000 + h64Lib, 2
            ReadMem64 VarPtr(p64Address), (tExportDirectory.AddressOfFunctions + lOrdinal * 4) / 10000 + h64Lib, 4
            
            GetProcAddress64 = p64Address + h64Lib
            
            Exit For
            
        End If
        
    Next

End Function

' //
' // Get 64-bit lib handle
' //
Public Property Get GetModuleHandle64( _
                    ByRef sLib As String) As Currency
    Dim tPBI64          As PROCESS_BASIC_INFORMATION64
    Dim lStatus         As Long
    Dim p64LdrData      As Currency
    Dim p64ListEntry    As Currency
    Dim p64LdrEntry     As Currency
    Dim p64DllName      As Currency
    Dim tusDll          As UNICODE_STRING64
    Dim tusLib          As UNICODE_STRING64

    GetMem4 StrPtr(sLib), tusDll.lpBuffer ' // Address
    tusDll.Length = LenB(sLib)
    tusDll.MaxLength = tusDll.Length + 2
    
    ' // We need 64-bit PEB
    lStatus = NtWow64QueryInformationProcess64(-1, ProcessBasicInformation, tPBI64, Len(tPBI64), 0)
    
    If lStatus < 0 Then
        Err.Raise lStatus
        Exit Property
    End If
    
    ' // Read PEB.Ldr
    ReadMem64 VarPtr(p64LdrData), tPBI64.PebBaseAddress + 0.0024@, Len(p64LdrData)
    
    p64ListEntry = p64LdrData + 0.0016@ ' // PEB_LDR_DATA.InLoadOrderModuleList.Flink
    
    ' // *PEB_LDR_DATA.InLoadOrderModuleList.Flink
    ReadMem64 VarPtr(p64LdrEntry), p64ListEntry, Len(p64LdrEntry)

    Do
        
        p64DllName = p64LdrEntry + 0.0088@ ' // LDR_DATA_TABLE_ENTRY.BaseDllName
        
        If Len(sLib) = 0 Then
            
            ReadMem64 VarPtr(GetModuleHandle64), p64LdrEntry + 0.0048@, Len(GetModuleHandle64)
            Exit Do
            
        Else
            
            ReadMem64 VarPtr(tusLib), p64DllName, Len(tusLib)
            
            If CompareUnicodeStrings64(tusLib, tusDll) = 0 Then
                
                ReadMem64 VarPtr(GetModuleHandle64), p64LdrEntry + 0.0048@, Len(GetModuleHandle64)
                Exit Do
                
            End If
        
        End If
        
        ReadMem64 VarPtr(p64LdrEntry), p64LdrEntry, Len(p64LdrEntry)

    Loop Until p64ListEntry = p64LdrEntry
    
End Property

' // Read memory at specified 64-bit address
Public Sub ReadMem64( _
           ByVal pTo As Long, _
           ByVal p64From As Currency, _
           ByVal lSize As Long)
    Dim lStatus As Long
    
    lStatus = NtWow64ReadVirtualMemory64(m_hCurHandle, p64From, ByVal pTo, lSize / 10000, 0)

    If lStatus < 0 Then
        Err.Raise lStatus
        Exit Sub
    End If
 
End Sub

' // Get null-terminated string length
Private Function StringLen64( _
                 ByVal p64 As Currency) As Currency
    Dim pAddrPair(1)    As Long
    Dim bPage()         As Byte
    Dim lSize           As Long
    Dim lStatus         As Long
    Dim lIndex          As Long
    Dim p64Start        As Currency
    
    p64Start = p64
    
    GetMem8 p64, pAddrPair(0)
    
    ' // Get number of bytes to end page boundry
    lSize = &H1000 - (pAddrPair(0) And &HFFF)
    
    Do
        
        ' // Read page
        ReDim Preserve bPage(lSize - 1)
        
        lStatus = NtWow64ReadVirtualMemory64(m_hCurHandle, p64, bPage(0), lSize / 10000, 0)
        
        If lStatus < 0 Then
            Err.Raise lStatus
            Exit Function
        End If
    
        For lIndex = 0 To lSize - 1
            
            ' // Test for null terminal
            If bPage(lIndex) = 0 Then
                
                StringLen64 = (p64 + lIndex / 10000) - p64Start
                Exit Do
                
            End If

        Next
        
        ' // Next page
        p64 = p64 + lSize / 10000
        
        lSize = 4096
        
    Loop While True
                 
End Function

' // Compare 2 ANSI strings
Private Function CompareAnsiStrings64( _
                 ByRef tasStr1 As ANSI_STRING64, _
                 ByRef tasStr2 As ANSI_STRING64, _
                 Optional ByVal bCaseSensitive As Boolean) As Long
    Dim bBuf1() As Byte
    Dim bBuf2() As Byte

    If tasStr1.Length > 0 Then
        
        ReDim bBuf1(tasStr1.Length)
        ReadMem64 VarPtr(bBuf1(0)), tasStr1.lpBuffer, tasStr1.Length
        
    End If
    
    If tasStr2.Length > 0 Then
    
        ReDim bBuf2(tasStr2.Length)
        ReadMem64 VarPtr(bBuf2(0)), tasStr2.lpBuffer, tasStr2.Length
        
    End If
    
    If bCaseSensitive Then
        CompareAnsiStrings64 = lstrcmp(bBuf1(0), bBuf2(0))
    Else
        CompareAnsiStrings64 = lstrcmpi(bBuf1(0), bBuf2(0))
    End If
    
End Function

' // Compare 2 strings
Private Function CompareUnicodeStrings64( _
                 ByRef tusStr1 As UNICODE_STRING64, _
                 ByRef tusStr2 As UNICODE_STRING64, _
                 Optional ByVal bCaseSensitive As Boolean) As Long
    Dim bBuf1() As Byte
    Dim bBuf2() As Byte

    If tusStr1.Length > 0 Then
        
        ReDim bBuf1(tusStr1.Length - 1)
        ReadMem64 VarPtr(bBuf1(0)), tusStr1.lpBuffer, tusStr1.Length
        
    End If
    
    If tusStr2.Length > 0 Then
    
        ReDim bBuf2(tusStr2.Length - 1)
        ReadMem64 VarPtr(bBuf2(0)), tusStr2.lpBuffer, tusStr2.Length
        
    End If

    If bCaseSensitive Then
        CompareUnicodeStrings64 = StrComp(bBuf1, bBuf2, vbBinaryCompare)
    Else
        CompareUnicodeStrings64 = StrComp(bBuf1, bBuf2, vbTextCompare)
    End If
    
End Function

