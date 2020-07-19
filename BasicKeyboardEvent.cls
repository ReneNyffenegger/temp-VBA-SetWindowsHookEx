' vi: ft=basic
'

option explicit

implements IKeyboardEvent

public function IKeyboardEvent_ev(vk_keyCode as long, down as boolean, alt as boolean) ' {

        debug.print "vk_keyCode = " & vk_keyCode & ", down = " & down & ", alt = " & alt

        IKeyboardEvent_ev = false

end function ' }
