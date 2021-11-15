Attribute VB_Name = "modscreenshot"
     Option Explicit
    Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal _
    bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    Public Const VK_SNAPSHOT As Byte = 44 ' PrintScreen virtual keycode
    Public Const PS_TheForm = 0
    Public Const PS_TheScreen As Byte = 1


