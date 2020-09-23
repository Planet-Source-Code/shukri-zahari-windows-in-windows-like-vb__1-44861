Attribute VB_Name = "modWinStyle"
' Module Name:   modWinStyle.bas
' Compatibility :   VB4-16, VB4-32, VB5, VB6
' Copyright      :   Copyright © 2002-2003, Shukri Zahari
' Comment      :  Make your own "designer" like VB does using this windows API functions
'                        have fun & don 't forget to vote me for this great code....
' Disclaimer     :  I, hereby don't hold any responsible for anything happen that could harm
'                        your PC or yourself!!! Use it at your own risk...

Option Explicit


#If Win16 Then          ' 16-bit declaration

    Declare Function FlashWindow% Lib "USER" (ByVal hWnd%, ByVal bInvert%)
    
    Declare Function SetWindowText% Lib "USER" (ByVal hWnd%, ByVal lpString$)
    
    Declare Function GetWindowLong& Lib "USER" (ByVal hWnd%, ByVal nIndex%)
    Declare Function SetWindowLong& Lib "USER" (ByVal hWnd%, ByVal nIndex%, ByVal dwNewLong&)
    
    Declare Sub ReleaseCapture Lib "USER" ()
    
    Declare Function SendMessage& Lib "USER" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lParam&)
    
#Else           ' 32-bit declaration
    
    Declare Function FlashWindow& Lib "USER32" (ByVal hWnd&, ByVal bInvert&)
    
    Declare Function SetWindowText& Lib "USER32" Alias "SetWindowTextA" (ByVal hWnd&, ByVal lpString$)
    
    Declare Function GetWindowLong& Lib "USER32" Alias "GetWindowLongA" (ByVal hWnd&, ByVal nIndex&)
    Declare Function SetWindowLong& Lib "USER32" Alias "SetWindowLongA" (ByVal hWnd&, ByVal nIndex&, ByVal dwNewLong&)
    
    Declare Sub ReleaseCapture Lib "USER32" ()
    
    Declare Function SendMessage& Lib "USER32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)

#End If


Global Const WM_SYSCOMMAND = &H112
Global Const MOUSE_MOVE = &HF012

Global Const GWL_STYLE = (-16)
Global Const GWL_EXSTYLE = (-20)

' ######################
' Windows Style
' For use with GWL_STYLE
'######################

Global Const WS_OVERLAPPED = &H0&
Global Const WS_POPUP = &H80000000
Global Const WS_CHILD = &H40000000
Global Const WS_MINIMIZE = &H20000000            ' Make the window "minimize"
Global Const WS_MINIMIZEBOX = &H20000           ' Show the minimize box
Global Const WS_VISIBLE = &H10000000              ' Show the window
Global Const WS_DISABLED = &H8000000             ' Disable the window
Global Const WS_CLIPSIBLINGS = &H4000000
Global Const WS_CLIPCHILDREN = &H2000000
Global Const WS_MAXIMIZE = &H1000000             ' Make the window "maximize"
Global Const WS_MAXIMIZEBOX = &H10000          ' Show the maximize box
Global Const WS_CAPTION = &HC00000                ' Set the window's caption
Global Const WS_BORDER = &H800000
Global Const WS_DLGFRAME = &H400000
Global Const WS_VSCROLL = &H200000               ' Add a Vertical scrollbar
Global Const WS_HSCROLL = &H100000               ' Add a Horizontal scrollbar
Global Const WS_SYSMENU = &H80000                ' Add system menu
Global Const WS_THICKFRAME = &H40000
Global Const WS_GROUP = &H20000
Global Const WS_TABSTOP = &H10000

Global Const WS_TILED = WS_OVERLAPPED
Global Const WS_ICONIC = WS_MINIMIZE             ' Same like make the window "minimize"
Global Const WS_SIZEBOX = WS_THICKFRAME

' ##########################
' Common Windows Style
' For use with GWL_STYLE
' ##########################

Global Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Global Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Global Const WS_CHILDWINDOW = (WS_CHILD)
Global Const WS_TILEDWINDOW = (WS_OVERLAPPEDWINDOW)

' #########################
' Extended Windows Style
' For use with GWL_STYLE
' #########################

Global Const WS_EX_DLGMODALFRAME = &H1&
Global Const WS_EX_NOPARENTNOTIFY = &H4&
Global Const WS_EX_TOPMOST = &H8&                     ' Set form on top of other form
Global Const WS_EX_ACCEPTFILES = &H10&              ' Maybe for drag-drop?
Global Const WS_EX_TRANSPARENT = &H20&            ' Transparent Window?

' #########################
' Dialog Style
' For use with GWL_STYLE
' #########################

Global Const DS_ABSALIGN = &H1&
Global Const DS_SYSMODAL = &H2&                     ' Make a modal form
Global Const DS_LOCALEDIT = &H20&
Global Const DS_SETFONT = &H40&
Global Const DS_MODALFRAME = &H80&
Global Const DS_NOIDLEMSG = &H100&


' #########################
' Button Style
' For use with GWL_STYLE
' #########################

' All are self-explain...

Global Const BS_PUSHBUTTON = &H0&
Global Const BS_DEFPUSHBUTTON = &H1&
Global Const BS_CHECKBOX = &H2&
Global Const BS_AUTOCHECKBOX = &H3&
Global Const BS_RADIOBUTTON = &H4&
Global Const BS_3STATE = &H5&
Global Const BS_AUTO3STATE = &H6&
Global Const BS_GROUPBOX = &H7&
Global Const BS_USERBUTTON = &H8&
Global Const BS_AUTORADIOBUTTON = &H9&
Global Const BS_PUSHBOX = &HA&
Global Const BS_OWNERDRAW = &HB&
Global Const BS_LEFTTEXT = &H20&


' #########################
' Listbox Style
' For use with GWL_Style
' #########################

Global Const LBS_NOTIFY = &H1&
Global Const LBS_SORT = &H2&
Global Const LBS_NOREDRAW = &H4&
Global Const LBS_MULTIPLESEL = &H8&
Global Const LBS_OWNERDRAWFIXED = &H10&
Global Const LBS_OWNERDRAWVARIABLE = &H20&
Global Const LBS_HASSTRINGS = &H40&
Global Const LBS_USETABSTOPS = &H80&
Global Const LBS_NOINTEGRALHEIGHT = &H100&
Global Const LBS_MULTICOLUMN = &H200&
Global Const LBS_WANTKEYBOARDINPUT = &H400&
Global Const LBS_EXTENDEDSEL = &H800&
Global Const LBS_DISABLENOSCROLL = &H1000&
Global Const LBS_STANDARD = (LBS_NOTIFY Or LBS_SORT Or WS_VSCROLL Or WS_BORDER)

' ########################
' Combobox Style
' For use with GWL_Style
'########################

Global Const CBS_SIMPLE = &H1&
Global Const CBS_DROPDOWN = &H2&
Global Const CBS_DROPDOWNLIST = &H3&
Global Const CBS_OWNERDRAWFIXED = &H10&
Global Const CBS_OWNERDRAWVARIABLE = &H20&
Global Const CBS_AUTOHSCROLL = &H40&
Global Const CBS_OEMCONVERT = &H80&
Global Const CBS_SORT = &H100&
Global Const CBS_HASSTRINGS = &H200&
Global Const CBS_NOINTEGRALHEIGHT = &H400&
Global Const CBS_DISABLENOSCROLL = &H800&

' ########################
' Scrollbox Style
' For use with GWL_Style
'########################

Global Const SBS_HORZ = &H0&
Global Const SBS_VERT = &H1&
Global Const SBS_TOPALIGN = &H2&
Global Const SBS_LEFTALIGN = &H2&
Global Const SBS_BOTTOMALIGN = &H4&
Global Const SBS_RIGHTALIGN = &H4&
Global Const SBS_SIZEBOXTOPLEFTALIGN = &H2&
Global Const SBS_SIZEBOXBOTTOMRIGHTALIGN = &H4&
Global Const SBS_SIZEBOX = &H8&

' ########################
' EDIT control Style
' For use with GWL_Style
' ########################

Global Const ES_LEFT = &H0&
Global Const ES_CENTER = &H1&
Global Const ES_RIGHT = &H2&
Global Const ES_MULTILINE = &H4&
Global Const ES_UPPERCASE = &H8&
Global Const ES_LOWERCASE = &H10&
Global Const ES_PASSWORD = &H20&
Global Const ES_AUTOVSCROLL = &H40&
Global Const ES_AUTOHSCROLL = &H80&
Global Const ES_NOHIDESEL = &H100&
Global Const ES_OEMCONVERT = &H400&
Global Const ES_READONLY = &H800&
Global Const ES_WANTRETURN = &H1000&

' ######################
' Static Control Style
' For use with GWL_STYLE
' ######################

Global Const SS_LEFT = &H0&
Global Const SS_CENTER = &H1&
Global Const SS_RIGHT = &H2&
Global Const SS_ICON = &H3&
Global Const SS_BLACKRECT = &H4&
Global Const SS_GRAYRECT = &H5&
Global Const SS_WHITERECT = &H6&
Global Const SS_BLACKFRAME = &H7&
Global Const SS_GRAYFRAME = &H8&
Global Const SS_WHITEFRAME = &H9&
Global Const SS_USERITEM = &HA&
Global Const SS_SIMPLE = &HB&
Global Const SS_LEFTNOWORDWRAP = &HC&
Global Const SS_NOPREFIX = &H80&

Type POINTAPI
    X As Integer
    Y As Integer
End Type

