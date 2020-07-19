option explicit

public const WH_CBT         =  5
public const WH_KEYBOARD_LL = 13 ' Low level keyboard events (compare with WH_KEYBOARD)
public const WH_SHELL       = 10 ' Notification of shell events, such as creation of top level windows.

public const HC_ACTION               = 0

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
     flags       as long ' bit 4: alt key was pressed
     time        as long
     dwExtraInfo as long
end  type ' }

function startLowLevelKeyboardHook() as long ' {

'   call setHook(hookHandle, WH_KEYBOARD_LL, addressOf LowLevelKeyboardProc)

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
'   nCode
'
'     MSDN says
'
'          If nCode is less than zero, the hook procedure must return the
'          value returned by CallNextHookEx.
'
'          If nCode is greater than or equal to zero, it is highly
'          recommended that you call CallNextHookEx and return the value it returns;
'          otherwise, other applications that have installed WH_CALLWNDPROCRET hooks will
'          not receive hook notifications and may behave incorrectly as a result. If the
'          hook procedure does not call CallNextHookEx, the return value should be zero.
'
'   Return value
'
'     MSDN says:
'
'          If the hook procedure processed the message, it may return a nonzero
'          value to prevent the system from passing the message to the rest of
'          the hook chain or the target window procedure.
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

'Q     upOrDown = keyUpOrDown(wParam       )
'Q     char     = keyChar    (lParam.vkCode)
'Q 
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
