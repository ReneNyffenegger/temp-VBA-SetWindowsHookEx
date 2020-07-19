option explicit

public const WH_CBT         =  5
public const WH_KEYBOARD_LL = 13 ' Low level keyboard events (compare with WH_KEYBOARD)
public const WH_SHELL       = 10 ' Notification of shell events, such as creation of top level windows.

public const HC_ACTION               = 0

global keyboard_ev as IKeyboardEvent

'
' TODO: should probably use longPtr etc.
'
declare function GetModuleHandle     lib "kernel32"     alias "GetModuleHandleA"      ( _
     byVal lpModuleName as string) as long

  ' CallNextHookEx {
    declare function CallNextHookEx      lib "user32"                      ( _
         byVal hHook        as long, _
         byVal nCode        as long, _
         byVal wParam       as long, _
               lParam       as any ) as long
  ' }

declare function SetWindowsHookEx    lib "user32"       alias "SetWindowsHookExA" ( _
     byVal idHook     as long, _
     byVal lpfn       as long, _
     byVal hmod       as long, _
     byVal dwThreadId as long) as long

declare function UnhookWindowsHookEx  lib "user32"                                ( _
     byVal hHook      as long) as long


' declare function UnhookWinEvent      lib "user32.dll"                             ( _
'      byRef hWinEventHook as long) as long
' ' }

type KBDLLHOOKSTRUCT ' {
     vkCode      as long ' virtual key code in range 1 .. 254
     scanCode    as long ' hardware code
     flags       as long
          ' bit 4: if set -> alt key was pressed
          ' bit 7: transition state, 0 -> the key is pressed, 1 -> key is being released.
     time        as long
     dwExtraInfo as long
end  type ' }

public const VK_LBUTTON              = &h001 ' { Virtual keys
public const VK_RBUTTON              = &h002
public const VK_CANCEL               = &h003 ' Implemented as Ctrl-Break on most keyboards
public const VK_MBUTTON              = &h004
public const VK_XBUTTON1             = &h005
public const VK_XBUTTON2             = &h006
public const VK_BACK                 = &h008
public const VK_TAB                  = &h009
public const VK_CLEAR                = &h00c
public const VK_RETURN               = &h00d ' Enter
public const VK_SHIFT                = &h010
public const VK_CONTROL              = &h011
public const VK_MENU                 = &h012
public const VK_PAUSE                = &h013
public const VK_CAPITAL              = &h014
public const VK_KANA                 = &h015
public const VK_HANGUEL              = &h015
public const VK_HANGUL               = &h015
public const VK_JUNJA                = &h017
public const VK_FINAL                = &h018
public const VK_HANJA                = &h019
public const VK_KANJI                = &h019
public const VK_ESCAPE               = &h01b
public const VK_CONVERT              = &h01c
public const VK_NONCONVERT           = &h01d
public const VK_ACCEPT               = &h01e
public const VK_MODECHANGE           = &h01f
public const VK_SPACE                = &h020
public const VK_PRIOR                = &h021
public const VK_NEXT                 = &h022
public const VK_END                  = &h023
public const VK_HOME                 = &h024
public const VK_LEFT                 = &h025
public const VK_UP                   = &h026
public const VK_RIGHT                = &h027
public const VK_DOWN                 = &h028
public const VK_SELECT               = &h029
public const VK_PRINT                = &h02a
public const VK_EXECUTE              = &h02b
public const VK_SNAPSHOT             = &h02c
public const VK_INSERT               = &h02d
public const VK_DELETE               = &h02e
public const VK_HELP                 = &h02f
public const VK_LWIN                 = &h05b
public const VK_RWIN                 = &h05c
public const VK_APPS                 = &h05d
public const VK_SLEEP                = &h05f
public const VK_NUMPAD0              = &h060
public const VK_NUMPAD1              = &h061
public const VK_NUMPAD2              = &h062
public const VK_NUMPAD3              = &h063
public const VK_NUMPAD4              = &h064
public const VK_NUMPAD5              = &h065
public const VK_NUMPAD6              = &h066
public const VK_NUMPAD7              = &h067
public const VK_NUMPAD8              = &h068
public const VK_NUMPAD9              = &h069
public const VK_MULTIPLY             = &h06a
public const VK_ADD                  = &h06b
public const VK_SEPARATOR            = &h06c
public const VK_SUBTRACT             = &h06d
public const VK_DECIMAL              = &h06e
public const VK_DIVIDE               = &h06f
public const VK_F1                   = &h070
public const VK_F2                   = &h071
public const VK_F3                   = &h072
public const VK_F4                   = &h073
public const VK_F5                   = &h074
public const VK_F6                   = &h075
public const VK_F7                   = &h076
public const VK_F8                   = &h077
public const VK_F9                   = &h078
public const VK_F10                  = &h079
public const VK_F11                  = &h07a
public const VK_F12                  = &h07b
public const VK_F13                  = &h07c
public const VK_F14                  = &h07d
public const VK_F15                  = &h07e
public const VK_F16                  = &h07f
public const VK_F17                  = &h080
public const VK_F18                  = &h081
public const VK_F19                  = &h082
public const VK_F20                  = &h083
public const VK_F21                  = &h084
public const VK_F22                  = &h085
public const VK_F23                  = &h086
public const VK_F24                  = &h087
public const VK_NUMLOCK              = &h090
public const VK_SCROLL               = &h091
public const VK_LSHIFT               = &h0a0
public const VK_RSHIFT               = &h0a1
public const VK_LCONTROL             = &h0a2
public const VK_RCONTROL             = &h0a3
public const VK_LMENU                = &h0a4 ' This is apparently the left  "Alt" key
public const VK_RMENU                = &h0a5 ' This is apparently the right "Alt" key
public const VK_BROWSER_BACK         = &h0a6
public const VK_BROWSER_FORWARD      = &h0a7
public const VK_BROWSER_REFRESH      = &h0a8
public const VK_BROWSER_STOP         = &h0a9
public const VK_BROWSER_SEARCH       = &h0aa
public const VK_BROWSER_FAVORITES    = &h0ab
public const VK_BROWSER_HOME         = &h0ac
public const VK_VOLUME_MUTE          = &h0ad
public const VK_VOLUME_DOWN          = &h0ae
public const VK_VOLUME_UP            = &h0af
public const VK_MEDIA_NEXT_TRACK     = &h0b0
public const VK_MEDIA_PREV_TRACK     = &h0b1
public const VK_MEDIA_STOP           = &h0b2
public const VK_MEDIA_PLAY_PAUSE     = &h0b3
public const VK_LAUNCH_MAIL          = &h0b4
public const VK_LAUNCH_MEDIA_SELECT  = &h0b5
public const VK_LAUNCH_APP1          = &h0b6
public const VK_LAUNCH_APP2          = &h0b7
public const VK_OEM_1                = &h0ba
public const VK_OEM_PLUS             = &h0bb
public const VK_OEM_COMMA            = &h0bc
public const VK_OEM_MINUS            = &h0bd
public const VK_OEM_PERIOD           = &h0be
public const VK_OEM_2                = &h0bf
public const VK_OEM_3                = &h0c0
public const VK_OEM_4                = &h0db
public const VK_OEM_5                = &h0dc
public const VK_OEM_6                = &h0dd
public const VK_OEM_7                = &h0de
public const VK_OEM_8                = &h0df
public const VK_OEM_102              = &h0e2
public const VK_PROCESSKEY           = &h0e5
public const VK_PACKET               = &h0e7
public const VK_ATTN                 = &h0f6
public const VK_EXSEL                = &h0f8
public const VK_PLAY                 = &h0fa
public const VK_NONAME               = &h0fc ' }

function startLowLevelKeyboardHook() as long ' {

'   call setHook(hookHandle, WH_KEYBOARD_LL, addressOf LowLevelKeyboardProc)

    set keyboard_ev = new BasicKeyboardEvent

    dim callBack as long ' TODO: as longPtr?

    callback = getAddressOfCallback(addressOf LowLevelKeyboardProc)

    startLowLevelKeyboardHook = SetWindowsHookEx( _
         WH_KEYBOARD_LL                         , _
         callBack                               , _
         GetModuleHandle(vbNullString)          , _
         0 )

    debug.print "Hook started, startLowLevelKeyboardHook = " & startLowLevelKeyboardHook

end function ' }

function endKeyboardHook(byVal hh as long) as boolean ' {

'   if hh <> 0 then
       endKeyboardHook = UnhookWindowsHookEx(hh)
       debug.print "endKeyboardHook = " & endKeyboardHook
'      hh = 0
'      endKeyboardHook = true
'      debug.print 
'      exit function
'    end if

'    endKeyboardHook = false
end function ' }

function LowLevelKeyboardProc(byVal nCode as Long, byVal wParam as long, lParam as KBDLLHOOKSTRUCT) as long ' {
'
'
'   MSDN says (https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms644985(v=vs.85))
'
'     nCode
'        A code the hook procedure uses to determine how to process the
'        message.
'
'        If nCode is less than zero, the hook procedure must pass the
'        message to the CallNextHookEx function without further processing and
'        should return the value returned by CallNextHookEx.
'
'        If nCode == HC_ACTION, the wParam and lParam parameters contain
'        information about a keyboard message.
'
'     wParam
'        One of
'          -  WM_KEYDOWN,
'          -  WM_KEYUP
'          -  WM_SYSKEYDOWN
'          -  WM_SYSKEYUP.
'
'     lParam
'          A pointer to a KBDLLHOOKSTRUCT structure.
'
'
'     Return value
'
'          If the hook procedure processed the message, it may return a nonzero
'          value to prevent the system from passing the message to the rest of
'          the hook chain or the target window procedure.
'
'     -----------------------------------------------------------------------------
'
'     The hook procedure should process a message in less time than the data entry specified in the
'             LowLevelHooksTimeout
'     value in the following registry key under
'             HKEY_CURRENT_USER\Control Panel\Desktop
'

    dim upOrDown as string
'   dim altKey   as boolean
    dim char     as string

    dim keyEventString as string

    debug.print "nCode = " & ncode

    if nCode <> HC_ACTION then
       LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, byVal lParam)
       exit function
    end if

    if not kev.ev(vk_keyCode := lparam.vkCode, down := lparam.flags and 256, alt := lparam.flags and 32 ) then ' {
     '
     ' Event was not processed, pass it on:
     '
       LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, byVal lParam)
       exit function
    end if ' }

    LowLevelKeyboardProc = 1
    exit function

'Q     upOrDown = keyUpOrDown(wParam       )
'Q     char     = keyChar    (lParam.vkCode)

'Q     keyEventString = char & " " & upOrDown
'Q   '
'Q   ' Apparently, the 5th bit is set if an ALT key was involved:
'Q   '
'Q   ' altKey = lParam.flags and 32
'Q 
'Q     debug.print char & " " & upOrDown
'Q 
'Q     if isSendingInput then
'Q        debug.print keyEventString & " - Is sending input"
'Q        LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, byVal lParam)
'Q        exit function
'Q     end if
'Q     debug.print keyEventString
'Q 
'Q     if isEventEqual(0, VK_PAUSE, false) then
'Q        StopTaskAutomator
'Q     end if
'Q 
'Q '   if expectingCommand then
'Q '      debug.print "expecting command"
'Q '   else
'Q '      debug.print "not expecting command"
'Q '   end if
'Q 
'Q 
'Q 
'Q     if wParam = WM_KEYDOWN or wParam = WM_SYSKEYDOWN then
'Q        nextKeyEv.pressed = true
'Q     else
'Q        nextKeyEv.pressed = false
'Q     end if
'Q 
'Q     nextKeyEv.vkCode = lParam.vkCode
'Q 
'Q     call storeKeyEvent(nextKeyEv)
'Q 
'Q 
'Q     if     cmdInitSequence then
'Q            debug.print "starting new command"
'Q 
'Q            call Beep(440, 200)
'Q            expectingCommand = true
'Q            commandSoFar     = ""
'Q 
'Q          ' 2018-07-25
'Q          ' LowLevelKeyboardProc = 1
'Q            LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
'Q            exit function
'Q 
'Q     elseif expectingCommand then
'Q 
'Q            dim ev as keyEv
'Q            dim c  as string
'Q 
'Q            ev = getLastKeyEvent(0)
'Q 
'Q            if not ev.pressed then
'Q               if chr(ev.vkCode) >= "A" and chr(ev.vkCode) <= "Z" then
'Q                   commandSoFar = commandSoFar + chr(ev.vkCode)
'Q                   expectingCommand = checkCommand(commandSoFar)
'Q               else
'Q                   expectingCommand = false
'Q               end if
'Q            end if
'Q 
'Q            LowLevelKeyboardProc = 1
'Q            exit function
'Q 
'Q     end if
'Q '
'Q '  '
'Q '  ' Display what the user has pressed.
'Q '  '(Needs Excel)
'Q '  '
'Q '    cells(1,1) = upOrDown
'Q '    cells(1,2) = lParam.vkCode
'Q '    cells(1,3) = char
'Q '    cells(1,4) = lParam.flags
'Q '
'Q '    if altKey then cells(1,5) = "alt" else cells(1,5) = "-"
'Q 
'Q '   debug.print "Calling next LowLevelKeyboardProc"
    LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)

end function ' }

function getAddressOfCallback(addr as long) as long ' {
      '
      '  TODO: use longPtr instead of long?
      '
      '  https://renenyffenegger.ch/notes/development/languages/VBA/language/operators/addressOf
      '
         getAddressOfCallback = addr
end function ' }
