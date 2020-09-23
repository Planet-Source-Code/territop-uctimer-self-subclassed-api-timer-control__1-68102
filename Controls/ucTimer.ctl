VERSION 5.00
Begin VB.UserControl ucTimer 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ucTimer.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ucTimer.ctx":04E3
End
Attribute VB_Name = "ucTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+  File Description:
'       ucTimer - A Selfsubclassed API Timer Control
'
'   Product Name:
'       ucTimer.ctl
'
'   Compatability:
'       Widnows: 98, ME, NT, 2K, XP
'
'   Software Developed by:
'       Paul R. Territo, Ph.D
'
'   Based on the following On-Line Articles
'       (Paul Caton - Self-Subclassser)
'           http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'       (Paul R. Territo, Ph.D - ucMsgBox)
'           http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=67387&lngWId=1
'
'   Legal Copyright & Trademarks:
'       Copyright © 2007, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'       Trademark ™ 2007, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Advance Research Systems shall not be liable for
'       any incidental or consequential damages suffered by any use of this software.
'       This software is owned by Paul R. Territo, Ph.D and is sold for use as a
'       license in accordance with the terms of the License Agreement in the
'       accompanying the documentation.
'
'   Contact Information:
'       For Technical Assistance:
'       pwterrito@insightbb.com
'
'   Modification(s) History:
'
'       10Mar07 - Initial Usercontrol Build (Modified from ucMsgBox, see above)
'
'-  Notes:
'
'   Build Date & Time: 3/10/2007 12:46:09 PM
Const Major As Long = 1
Const Minor As Long = 0
Const Revision As Long = 0
Const DateTime As String = "3/10/2007 12:46:09 PM"
'
'   Force Declarations
Option Explicit

'   Private Constants
Private Const VER_PLATFORM_WIN32_NT As Long = 2

'   Private Class Priority Constants
Private Const REALTIME_PRIORITY_CLASS = &H100
Private Const HIGH_PRIORITY_CLASS = &H80
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const IDLE_PRIORITY_CLASS = &H40

'   Os Window Types
Private Type OSVERSIONINFO
    OSVSize         As Long         'size, in bytes, of this data structure
    dwVerMajor      As Long         'ie NT 3.51, dwVerMajor = 3; NT 4.0, dwVerMajor = 4.
    dwVerMinor      As Long         'ie NT 3.51, dwVerMinor = 51; NT 4.0, dwVerMinor= 0.
    dwBuildNumber   As Long         'NT: build number of the OS
                                    'Win9x: build number of the OS in low-order word.
                                    '       High-order word contains major & minor ver nos.
    PlatformID      As Long         'Identifies the operating system platform.
    szCSDVersion    As String * 128 'NT: string, such as "Service Pack 3"
                                    'Win9x: string providing arbitrary additional information
End Type

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long

Public Enum utPriorityEnum
    [utRealTime] = REALTIME_PRIORITY_CLASS          'Highest Priority
    [utHigh] = HIGH_PRIORITY_CLASS                  'Second Highest Priority
    [utNormal] = NORMAL_PRIORITY_CLASS              'Third Highest Priority
    [utIdle] = IDLE_PRIORITY_CLASS                  'Lowest Priority
End Enum
#If False Then
    Const utRealTime = REALTIME_PRIORITY_CLASS
    Const utHigh = HIGH_PRIORITY_CLASS
    Const utNormal = NORMAL_PRIORITY_CLASS
    Const utIdle = IDLE_PRIORITY_CLASS
#End If

Public Enum utTimerTypeEnum
    [utCountUp] = &H0                               'Count Up Timer (0,1,2,3...n)
    [utCoundDown] = &H1                             'Count Down Timer (n, n-1, n-2...0)
End Enum
#If False Then
    Const utCountUp = &H0
    Const utCoundDown = &H1
#End If

'   Required Type Definitions
Private Type utTimerInfoType
    Duration As Long                                'How long a Timer Waits
    Elapsed As Long                                 'Total Time
    Enabled As Boolean                              'Enabled State of Timer
    ID As Long                                      'Unique ID for the Timer
    Interval As Long                                'Interval for the Ticks
    Remaining As Long                               'Time Left on the Clock
    ThreadPriority As utPriorityEnum                'Process Priority for MsgBox Timer
    TimerType As utTimerTypeEnum                    'Timer Type..CountUp/Down
End Type

'   Private variables
Private bTimerRunning       As Boolean              'Timer Running Flag for SelfClose
Private m_Timer             As utTimerInfoType      'Timer info structure

Public Event Timer()
Public Event Initailized()
Public Event Elapsed(nTime As Long)
Public Event Remaining(nTime As Long)
Public Event Terminated()

'==================================================================================================
' ucSubclass - A template UserControl for control authors that require self-subclassing without ANY
'              external dependencies. IDE safe.
'
' Paul_Caton@hotmail.com
' Copyright free, use and abuse as you see fit.
'
' v1.0.0000 20040525 First cut.....................................................................
' v1.1.0000 20040602 Multi-subclassing version.....................................................
' v1.1.0001 20040604 Optimized the subclass code...................................................
' v1.1.0002 20040607 Substituted byte arrays for strings for the code buffers......................
' v1.1.0003 20040618 Re-patch when adding extra hWnds..............................................
' v1.1.0004 20040619 Optimized to death version....................................................
' v1.1.0005 20040620 Use allocated memory for code buffers, no need to re-patch....................
' v1.1.0006 20040628 Better protection in zIdx, improved comments..................................
' v1.1.0007 20040629 Fixed InIDE patching oops.....................................................
' v1.1.0008 20040910 Fixed bug in UserControl_Terminate, zSubclass_Proc procedure hidden...........
'==================================================================================================
'Subclasser declarations

Public Event Status(ByVal sStatus As String)

Private Const WM_TIMER                  As Long = &H113

Private bInCtrl                      As Boolean
Private bSubClass                    As Boolean

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

Private Enum eMsgWhen
    MSG_AFTER = 1                                                                   'Message calls back after the original (previous) WndProc
    MSG_BEFORE = 2                                                                  'Message calls back before the original (previous) WndProc
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                  'Message calls back before and after the original (previous) WndProc
End Enum
#If False Then
    Private Const MSG_AFTER = 1                                                                   'Message calls back after the original (previous) WndProc
    Private Const MSG_BEFORE = 2                                                                  'Message calls back before the original (previous) WndProc
    Private Const MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                  'Message calls back before and after the original (previous) WndProc
#End If

Private Const ALL_MESSAGES           As Long = -1                                   'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                    'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                   'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                   'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                   'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                  'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                  'Table A (after) entry count patch offset

Private Type tSubData                                                               'Subclass data type
    hWnd                               As Long                                      'Handle of the window being subclassed
    nAddrSub                           As Long                                      'The address of our new WndProc (allocated memory).
    nAddrOrig                          As Long                                      'The address of the pre-existing WndProc
    nMsgCntA                           As Long                                      'Msg after table entry count
    nMsgCntB                           As Long                                      'Msg before table entry count
    aMsgTblA()                         As Long                                      'Msg after table array
    aMsgTblB()                         As Long                                      'Msg Before table array
End Type

Private sc_aSubData()                As tSubData                                    'Subclass data array

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)

'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
Attribute zSubclass_Proc.VB_MemberFlags = "40"
    'Parameters:
        'bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
        'bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
        'lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
        'hWnd     - The window handle
        'uMsg     - The message number
        'wParam   - Message related data
        'lParam   - Message related data
    'Notes:
        'If you really know what you're doing, it's possible to change the values of the
        'hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
        'values get passed to the default handler.. and optionaly, the 'after' callback
    
    Select Case uMsg
        Case WM_TIMER
            Call TimerProc
                        
    End Select
    
End Sub

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

    'Add a message to the table of those that will invoke a callback. You should Subclass_Subclass first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    'Parameters:
        'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
        'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
        'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    'Parameters:
    'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
    'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
    'When      - Whether the msg is to be removed from the before, after or both callback tables
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
    '   Determine if the passed function is supported
                                    
    On Error GoTo IsFunctionExported_Error
    
    Dim hmod        As Long
    Dim bLibLoaded  As Boolean
    
    hmod = GetModuleHandleA(sModule)
    
    If hmod = 0 Then
        hmod = LoadLibraryA(sModule)
        If hmod Then
            bLibLoaded = True
        End If
    End If
    
    If hmod Then
        If GetProcAddress(hmod, sFunction) Then
            IsFunctionExported = True
        End If
    End If
    
    If bLibLoaded Then
        Call FreeLibrary(hmod)
    End If
    
    Exit Function

IsFunctionExported_Error:
End Function

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
    'Parameters:
    'lng_hWnd  - The handle of the window to be subclassed
    'Returns;
    'The sc_aSubData() index
    Const CODE_LEN              As Long = 204                                       'Length of the machine code in bytes
    Const FUNC_CWP              As String = "CallWindowProcA"                       'We use CallWindowProc to call the original WndProc
    Const FUNC_EBM              As String = "EbMode"                                'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
    Const FUNC_SWL              As String = "SetWindowLongA"                        'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
    Const MOD_USER              As String = "user32"                                'Location of the SetWindowLongA & CallWindowProc functions
    Const MOD_VBA5              As String = "vba5"                                  'Location of the EbMode function if running VB5
    Const MOD_VBA6              As String = "vba6"                                  'Location of the EbMode function if running VB6
    Const PATCH_01              As Long = 18                                        'Code buffer offset to the location of the relative address to EbMode
    Const PATCH_02              As Long = 68                                        'Address of the previous WndProc
    Const PATCH_03              As Long = 78                                        'Relative address of SetWindowsLong
    Const PATCH_06              As Long = 116                                       'Address of the previous WndProc
    Const PATCH_07              As Long = 121                                       'Relative address of CallWindowProc
    Const PATCH_0A              As Long = 186                                       'Address of the owner object
    Static aBuf(1 To CODE_LEN)  As Byte                                             'Static code buffer byte array
    Static pCWP                 As Long                                             'Address of the CallWindowsProc
    Static pEbMode              As Long                                             'Address of the EbMode IDE break/stop/running function
    Static pSWL                 As Long                                             'Address of the SetWindowsLong function
    Dim i                       As Long                                             'Loop index
    Dim j                       As Long                                             'Loop index
    Dim nSubIdx                 As Long                                             'Subclass data index
    Dim sHex                    As String                                           'Hex code string
    
    'If it's the first time through here..
    If aBuf(1) = 0 Then
        
        'The hex pair machine code representation.
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
            "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
            "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
            "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
        
        'Convert the string from hex pairs to bytes and store in the static machine code buffer
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                                  'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop                                                                        'Next pair of hex characters
        
        'Get API function addresses
        If Subclass_InIDE Then                                                      'If we're running in the VB IDE
            aBuf(16) = &H90                                                         'Patch the code buffer to enable the IDE state code
            aBuf(17) = &H90                                                         'Patch the code buffer to enable the IDE state code
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                 'Get the address of EbMode in vba6.dll
            If pEbMode = 0 Then                                                     'Found?
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                             'VB5 perhaps
            End If
        End If

        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                        'Get the address of the CallWindowsProc function
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                        'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData                                       'Create the first sc_aSubData element
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If nSubIdx = -1 Then                                                        'If an sc_aSubData element isn't being re-cycled
            nSubIdx = UBound(sc_aSubData()) + 1                                     'Calculate the next element
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                    'Create a new sc_aSubData element
        End If
        Subclass_Start = nSubIdx
    End If
    
    '   Use the following debuging to indicate which index into the
    '   sc_aSubData array the AddressOf Pointer exists at....
    'Debug.Print "AddressOf Index: " & nSubIdx
    With sc_aSubData(nSubIdx)
        .hWnd = lng_hWnd                                                            'Store the hWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                               'Allocate memory for the machine code WndProc
        .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                  'Set our WndProc in place
        Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                      'Copy the machine code from the static byte array to the code array in sc_aSubData
        Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                             'Original WndProc address for CallWindowProc, call the original WndProc
        Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                   'Patch the relative address of the SetWindowLongA api function
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                             'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                   'Patch the relative address of the CallWindowProc api function
        Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                             'Patch the address of this object instance into the static machine code buffer
    End With
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
    Dim i As Long
    
    i = UBound(sc_aSubData())                                                       'Get the upper bound of the subclass data array
    Do While i >= 0                                                                 'Iterate through each element
        With sc_aSubData(i)
            If .hWnd <> 0 Then                                                      'If not previously Subclass_Stop'd
                Call Subclass_Stop(.hWnd)                                           'Subclass_Stop
            End If
        End With
        i = i - 1                                                                   'Next element
    Loop
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
    'Parameters:
    'lng_hWnd  - The handle of the window to stop being subclassed
    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)                         'Restore the original WndProc
        Call zPatchVal(.nAddrSub, PATCH_05, 0)                                      'Patch the Table B entry count to ensure no further 'before' callbacks
        Call zPatchVal(.nAddrSub, PATCH_09, 0)                                      'Patch the Table A entry count to ensure no further 'after' callbacks
        Call GlobalFree(.nAddrSub)                                                  'Release the machine code memory
        .hWnd = 0                                                                   'Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0                                                               'Clear the before table
        .nMsgCntA = 0                                                               'Clear the after table
        Erase .aMsgTblB                                                             'Erase the before table
        Erase .aMsgTblA                                                             'Erase the after table
    End With
End Sub

Private Function TimerProc()
    
    On Error Resume Next
    
    'Debug.Print "Running..." & Timer()
    If m_Timer.Enabled Then
        RaiseEvent Timer
        Select Case m_Timer.TimerType
            Case utCountUp
                '   Normal
                m_Timer.Elapsed = m_Timer.Elapsed + m_Timer.Interval
                RaiseEvent Elapsed(m_Timer.Elapsed)
            Case utCoundDown
                '   This requires the duration to be set as well!
                m_Timer.Elapsed = m_Timer.Elapsed + m_Timer.Interval
                RaiseEvent Remaining(m_Timer.Duration - (m_Timer.Elapsed))
                If (m_Timer.Duration - (m_Timer.Elapsed)) <= 0 Then
                    Call pvSetTimer(False)
                End If
        End Select
    Else
        Call pvSetTimer(m_Timer.Enabled)
    End If

End Function

'======================================================================================================
'These z??? routines are exclusively called by the Subclass_??? routines.

'Worker sub for sc_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry  As Long                                                             'Message table entry index
    Dim nOff1   As Long                                                             'Machine code buffer offset 1
    Dim nOff2   As Long                                                             'Machine code buffer offset 2
    
    If uMsg = ALL_MESSAGES Then                                                     'If all messages
        nMsgCnt = ALL_MESSAGES                                                      'Indicates that all messages will callback
    Else                                                                            'Else a specific message number
        Do While nEntry < nMsgCnt                                                   'For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1
            
            If aMsgTbl(nEntry) = 0 Then                                             'This msg table slot is a deleted entry
                aMsgTbl(nEntry) = uMsg                                              'Re-use this entry
                Exit Sub                                                            'Bail
            ElseIf aMsgTbl(nEntry) = uMsg Then                                      'The msg is already in the table!
                Exit Sub                                                            'Bail
            End If
        Loop                                                                        'Next entry
        nMsgCnt = nMsgCnt + 1                                                       'New slot required, bump the table entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                'Bump the size of the table.
        aMsgTbl(nMsgCnt) = uMsg                                                     'Store the message number in the table
    End If

    If When = eMsgWhen.MSG_BEFORE Then                                              'If before
        nOff1 = PATCH_04                                                            'Offset to the Before table
        nOff2 = PATCH_05                                                            'Offset to the Before table entry count
    Else                                                                            'Else after
        nOff1 = PATCH_08                                                            'Offset to the After table
        nOff2 = PATCH_09                                                            'Offset to the After table entry count
    End If

    If uMsg <> ALL_MESSAGES Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                            'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)                                           'Patch the appropriate table entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc                                                          'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for sc_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry As Long
    
    If uMsg = ALL_MESSAGES Then                                                     'If deleting all messages
        nMsgCnt = 0                                                                 'Message count is now zero
        If When = eMsgWhen.MSG_BEFORE Then                                          'If before
            nEntry = PATCH_05                                                       'Patch the before table message count location
        Else                                                                        'Else after
            nEntry = PATCH_09                                                       'Patch the after table message count location
        End If
        Call zPatchVal(nAddr, nEntry, 0)                                            'Patch the table message count to zero
    Else                                                                            'Else deleteting a specific message
        Do While nEntry < nMsgCnt                                                   'For each table entry
            nEntry = nEntry + 1
            If aMsgTbl(nEntry) = uMsg Then                                          'If this entry is the message we wish to delete
                aMsgTbl(nEntry) = 0                                                 'Mark the table slot as available
                Exit Do                                                             'Bail
            End If
        Loop                                                                        'Next entry
    End If
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
    'Get the upper bound of sc_aSubData() - If you get an error here, you're probably sc_AddMsg-ing before Subclass_Start
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0                                                              'Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)
            If .hWnd = lng_hWnd Then                                                'If the hWnd of this element is the one we're looking for
                If Not bAdd Then                                                    'If we're searching not adding
                    Exit Function                                                   'Found
                End If
            ElseIf .hWnd = 0 Then                                                   'If this an element marked for reuse.
                If bAdd Then                                                        'If we're adding
                    Exit Function                                                   'Re-use it
                End If
            End If
        End With
    zIdx = zIdx - 1                                                                 'Decrement the index
    Loop
    
    If Not bAdd Then
        '   Never, Ever use this in a modal dialog or your system will hang!!!
        'Debug.Assert False     'hWnd not found, programmer error
        '   Instead, we need a way to get out gracefully, so stop everything
        '   and continue on processing the requests....
        Call Subclass_StopAll
        Debug.Print "Sublcassing Error....No Handle Located!!!"
    End If

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function
'======================================================================================================
'   End SubClass Sections
'======================================================================================================

Public Property Get BackColor() As OLE_COLOR
    
    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler
    
    BackColor = UserControl.BackColor
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucTimer.BackColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    
    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler
    
    UserControl.BackColor = NewValue
    PropertyChanged "BackColor"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucTimer.BackColor", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get Duration() As Long

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Duration = m_Timer.Duration
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucTimer.Duration", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let Duration(ByVal NewValue As Long)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_Timer.Duration = NewValue
    PropertyChanged "Duration"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucTimer.Duration", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get Enabled() As Boolean

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Enabled = m_Timer.Enabled
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucTimer.Duration", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_Timer.Enabled = NewValue
    Call pvSetTimer(m_Timer.Enabled)
    PropertyChanged "Enabled"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucTimer.Duration", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Get Interval() As Long
Attribute Interval.VB_UserMemId = 0
Attribute Interval.VB_MemberFlags = "200"

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    Interval = m_Timer.Interval
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucTimer.Interval", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let Interval(ByVal NewValue As Long)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_Timer.Interval = NewValue
    PropertyChanged "Interval"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucTimer.Interval", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Function IsWinXP() As Boolean
    'returns True if running Windows XP
    Dim OSV As OSVERSIONINFO

    '   Handle Any Errors
    On Error GoTo Func_ErrHandler

    OSV.OSVSize = Len(OSV)
    If GetVersionEx(OSV) = 1 Then
        IsWinXP = (OSV.PlatformID = VER_PLATFORM_WIN32_NT) And _
            (OSV.dwVerMajor = 5 And OSV.dwVerMinor = 1) And _
            (OSV.dwBuildNumber >= 2600)
    End If
    
Func_ErrHandlerExit:
    Exit Function
Func_ErrHandler:
    Err.Raise Err.Number, "ucTimer.IsWinXP", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Func_ErrHandlerExit:
End Function

Private Sub SetProcessPriority(ByVal lPriority As utPriorityEnum)
    Dim lRet As Long
    Dim lProcessID As Long
    Dim lProcessHandle As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    '   Get the Process Handled
    lProcessHandle = GetCurrentProcess
    'Debug.Print "lProcessHandle: " & lProcessHandle
    '   Sets the priority using any priority from the following
    '   Highest priority to the Lowest ;-)
    '   - REALTIME_PRIORITY_CLASS
    '   - HIGH_PRIORITY_CLASS
    '   - NORMAL_PRIORITY_CLASS
    '   - IDLE_PRIORITY_CLASS
    '   If lRet <> 0 then the changing operation action was successful
    lRet = SetPriorityClass(lProcessHandle, lPriority)

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucTimer.SetProcessPriority", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub pvSetTimer(ByVal bEnabled As Boolean)
    
    On Error Resume Next
    If Not bEnabled Then
        With m_Timer
            .ID = KillTimer(UserControl.hWnd, .ID)
            .Elapsed = 0
            .Remaining = .Duration
        End With
        RaiseEvent Terminated
        bTimerRunning = False
    Else
        If Not bTimerRunning Then
            With m_Timer
                .ID = SetTimer(UserControl.hWnd, .ID, .Interval, sc_aSubData(0).nAddrSub)
            End With
            RaiseEvent Initailized
            bTimerRunning = True
        End If
    End If

End Sub

Public Property Get ThreadPriority() As utPriorityEnum

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    ThreadPriority = m_Timer.ThreadPriority
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucTimer.ThreadPriority", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let ThreadPriority(ByVal NewValue As utPriorityEnum)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    m_Timer.ThreadPriority = NewValue
    PropertyChanged "ThreadPriority"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucTimer.ThreadPriority", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

'TimerType
Public Property Get TimerType() As utTimerTypeEnum

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    TimerType = m_Timer.TimerType
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucTimer.ThreadPriority", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Public Property Let TimerType(ByVal NewValue As utTimerTypeEnum)

    '   Handle Any Errors
    On Error GoTo Prop_ErrHandler

    If NewValue <> m_Timer.TimerType Then
        m_Timer.TimerType = NewValue
        If bTimerRunning Then
            Call pvSetTimer(False)
            Enabled = True
        End If
    End If
    PropertyChanged "ThreadPriority"
    
Prop_ErrHandlerExit:
    Exit Property
Prop_ErrHandler:
    Err.Raise Err.Number, "ucTimer.ThreadPriority", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Prop_ErrHandlerExit:
End Property

Private Sub UserControl_InitProperties()
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    With m_Timer
        .Duration = 1000                'Miliseconds
        .Enabled = False                'Disabled
        .ID = 0                         'Use none to start
        .Interval = 10                  'Miliseconds
        .ThreadPriority = utRealTime    'Realtime Thread
        .TimerType = utCountUp          'Count Up
    End With
    UserControl.BackColor = UserControl.Parent.BackColor

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucTimer.UserControl_InitProperties", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    Dim i As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    With PropBag
        m_Timer.Duration = .ReadProperty("Duration", 1000)
        m_Timer.Enabled = .ReadProperty("Enabled", False)
        m_Timer.ID = .ReadProperty("ID", 0&)
        m_Timer.Interval = .ReadProperty("Interval", 10)
        m_Timer.ThreadPriority = .ReadProperty("ThreadPriority", [utRealTime])
        m_Timer.TimerType = .ReadProperty("TimerType", [utCountUp])
    End With
    
    If Ambient.UserMode Then 'If we're not in design mode
        'OS supports mouse leave so subclass for it
        With UserControl
            'Start subclassing the UserControl
            m_Timer.ID = Subclass_Start(UserControl.hWnd)
            Call Subclass_AddMsg(UserControl.hWnd, WM_TIMER, MSG_AFTER)
            RaiseEvent Status("Initailizing Subclassing")
        End With
    End If
    UserControl.BackColor = UserControl.Parent.BackColor
    UserControl_Resize

Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucTimer.UserControl_ReadProperties", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_Resize()
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler
    
    With UserControl
        .Width = 375
        .Height = 375
    End With
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucTimer.UserControl_Resize", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_Show()
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    UserControl.BackColor = UserControl.Parent.BackColor
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucTimer.UserControl_Show", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Private Sub UserControl_Terminate()
    '   The control is terminating - a good place to stop the subclasser
    '
    '   We will keep this code in place...just in case,
    '   as the Host Objects QueryUnload should have Stopped All
    '   subclassing before we got here....
    On Error GoTo Catch
    If bSubClass Then
        Call pvSetTimer(False)
        Call Subclass_StopAll
        RaiseEvent Status("Terminating Subclassing")
    End If
Catch:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim i As Long
    
    '   Handle Any Errors
    On Error GoTo Sub_ErrHandler

    With PropBag
        Call .WriteProperty("Duration", m_Timer.Duration, 1000)
        Call .WriteProperty("Enabled", m_Timer.Enabled, False)
        Call .WriteProperty("ID", m_Timer.ID, 0&)
        Call .WriteProperty("Interval", m_Timer.Interval, 10)
        Call .WriteProperty("ThreadPriority", m_Timer.ThreadPriority, [utRealTime])
        Call .WriteProperty("TimerType", m_Timer.TimerType, [utCountUp])
    End With
    
Sub_ErrHandlerExit:
    Exit Sub
Sub_ErrHandler:
    Err.Raise Err.Number, "ucTimer.UserControl_WriteProperties", Err.Description, Err.HelpFile, Err.HelpContext
    Resume Sub_ErrHandlerExit:
End Sub

Public Function Version(Optional ByVal bDateTime As Boolean) As String
    On Error GoTo Version_Error
    
    If bDateTime Then
        Version = Major & "." & Minor & "." & Revision & " (" & DateTime & ")"
    Else
        Version = Major & "." & Minor & "." & Revision
    End If
    Exit Function
    
Version_Error:
End Function





