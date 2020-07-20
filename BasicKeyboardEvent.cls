' vi: ft=basic
'

option explicit

implements IKeyboardEvent

public function IKeyboardEvent_ev(vk_keyCode as long, pressed as boolean, alt as boolean, scanCode as long, time as long) ' {

        debug.print "vk_keyCode = " & vk_keyCode & ", pressed = " & pressed & ", alt = " & alt & ", scanCode = " & scanCode & ", time = " & time

        IKeyboardEvent_ev = false

end function ' }
