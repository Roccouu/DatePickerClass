VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatePickerClass_v201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ================================ DATE PICKER CLASS ================================='
' =============================== CREDITS AND LICENSE ================================'
' ------------------------------------------------------------------------------------'
' MIT License
' ------------------------------------------------------------------------------------'
' Copyright (c) 2021 Roberto Carlos Romay Medina (Roccou)
' DatePickerClass, Excel simple date picker for use with macros.
' Version: 2.0.1
' GitHub: https://github.com/roccouu/
' Potosí - Bolívia
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
' ------------------------------------------------------------------------------------'
' ACKNOWLEDGMENS
' ------------------------------------------------------------------------------------'
' INSPIRED ON ANDRÉS ROJAS MONCADA'S CALENDAR CORE ALGORITHM - Thanks a lot AndrÃ©s for all your great talent.
' Nota del autor ---------------------------------------------------------------------'
' Andrés Rojas Moncada - Autor del canal Excel Hecho Fácil en YouTube
' Versión 1.0 - 20 de julio de 2015
' URL del canal: www.youtube.com/jarmoncada01
' Para ver algunos ejemplos sobre el uso de este calendario, observa este video.
' Enlace: |||||||||||||||||||| https://www.youtube.com/watch?v=FkjsuN2zqSU ||||||||||||||||||||
' ! Muchas gracias y espero lo disfruten ! -------------------------------------------'
'
' =============================== CREDITS AND LICENSE ================================'
' ================================ DATE PICKER CLASS ================================='



' =========================== DATE PICKER CLASS STRUCTURE ==========================='
' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
' VARIABLES (GLOStr_) ---------------------------------------------------------------'
' CONSTANTS (GLOSTR_) ---------------------------------------------------------------'
' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
' PROPERTIES ------------------------------------------------------------------------'
' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
' CONTROLS EVENTS -------------------------------------------------------------------'
' CORE METHODS LIST PRIVATE ---------------------------------------------------------'
' DatePicker_open
' DatePicker_close
' DatePicker_controls_event_select_date
' DatePicker_controls_event_mouse_move
' DatePicker_core_InicializaFormularioCalendario
' DatePicker_core_CargarLosDias
' DatePicker_core_change_month_year
' DatePicker_core_remark_the_day
' DatePicker_core_remark_selected_day
' DatePicker_core_change_colors
' DatePicker_core_add
' DatePicker_core_use
' DatePicker_core_set_controls
' DatePicker_core_reset
' ASSETS ----------------------------------------------------------------------------'
' Datepicker_asset_get_horizontal_align
' Datepicker_asset_get_vertical_align
' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
' Debug_assert
' PUBLIC METHODS LIST (INTERFACE) ---------------------------------------------------'
' DatePicker
' =========================== DATE PICKER CLASS STRUCTURE ==========================='



' =========================== DATE PICKER CLASS STRUCTURE ============================'
' OBJECT VARIABLES (GLOOBJ_) --------------------------------------------------------'
Private GLOOBJ_FraDatepicker1 As MSForms.Frame ' The DPC main object
Private GLOOBJ_FraDatepicker2 As MSForms.Frame
Private GLOOBJ_CtlReceptor As Object 'MSForms.Control

Private WithEvents GLOOBJ_LblHead2 As MSForms.Label
Attribute GLOOBJ_LblHead2.VB_VarHelpID = -1

Private WithEvents GLOOBJ_SpinHead1 As MSForms.SpinButton
Attribute GLOOBJ_SpinHead1.VB_VarHelpID = -1
Private WithEvents GLOOBJ_SpinHead2 As MSForms.SpinButton
Attribute GLOOBJ_SpinHead2.VB_VarHelpID = -1

Private WithEvents GLOOBJ_LblBody1 As MSForms.Label
Attribute GLOOBJ_LblBody1.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody2 As MSForms.Label
Attribute GLOOBJ_LblBody2.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody3 As MSForms.Label
Attribute GLOOBJ_LblBody3.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody4 As MSForms.Label
Attribute GLOOBJ_LblBody4.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody5 As MSForms.Label
Attribute GLOOBJ_LblBody5.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody6 As MSForms.Label
Attribute GLOOBJ_LblBody6.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody7 As MSForms.Label
Attribute GLOOBJ_LblBody7.VB_VarHelpID = -1

Private WithEvents GLOOBJ_LblBody8 As MSForms.Label
Attribute GLOOBJ_LblBody8.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody9 As MSForms.Label
Attribute GLOOBJ_LblBody9.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody10 As MSForms.Label
Attribute GLOOBJ_LblBody10.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody11 As MSForms.Label
Attribute GLOOBJ_LblBody11.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody12 As MSForms.Label
Attribute GLOOBJ_LblBody12.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody13 As MSForms.Label
Attribute GLOOBJ_LblBody13.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody14 As MSForms.Label
Attribute GLOOBJ_LblBody14.VB_VarHelpID = -1

Private WithEvents GLOOBJ_LblBody15 As MSForms.Label
Attribute GLOOBJ_LblBody15.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody16 As MSForms.Label
Attribute GLOOBJ_LblBody16.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody17 As MSForms.Label
Attribute GLOOBJ_LblBody17.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody18 As MSForms.Label
Attribute GLOOBJ_LblBody18.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody19 As MSForms.Label
Attribute GLOOBJ_LblBody19.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody20 As MSForms.Label
Attribute GLOOBJ_LblBody20.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody21 As MSForms.Label
Attribute GLOOBJ_LblBody21.VB_VarHelpID = -1

Private WithEvents GLOOBJ_LblBody22 As MSForms.Label
Attribute GLOOBJ_LblBody22.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody23 As MSForms.Label
Attribute GLOOBJ_LblBody23.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody24 As MSForms.Label
Attribute GLOOBJ_LblBody24.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody25 As MSForms.Label
Attribute GLOOBJ_LblBody25.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody26 As MSForms.Label
Attribute GLOOBJ_LblBody26.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody27 As MSForms.Label
Attribute GLOOBJ_LblBody27.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody28 As MSForms.Label
Attribute GLOOBJ_LblBody28.VB_VarHelpID = -1

Private WithEvents GLOOBJ_LblBody29 As MSForms.Label
Attribute GLOOBJ_LblBody29.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody30 As MSForms.Label
Attribute GLOOBJ_LblBody30.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody31 As MSForms.Label
Attribute GLOOBJ_LblBody31.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody32 As MSForms.Label
Attribute GLOOBJ_LblBody32.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody33 As MSForms.Label
Attribute GLOOBJ_LblBody33.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody34 As MSForms.Label
Attribute GLOOBJ_LblBody34.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody35 As MSForms.Label
Attribute GLOOBJ_LblBody35.VB_VarHelpID = -1

Private WithEvents GLOOBJ_LblBody36 As MSForms.Label
Attribute GLOOBJ_LblBody36.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody37 As MSForms.Label
Attribute GLOOBJ_LblBody37.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody38 As MSForms.Label
Attribute GLOOBJ_LblBody38.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody39 As MSForms.Label
Attribute GLOOBJ_LblBody39.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody40 As MSForms.Label
Attribute GLOOBJ_LblBody40.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody41 As MSForms.Label
Attribute GLOOBJ_LblBody41.VB_VarHelpID = -1
Private WithEvents GLOOBJ_LblBody42 As MSForms.Label
Attribute GLOOBJ_LblBody42.VB_VarHelpID = -1

' VARIABLES (GLOStr_) ---------------------------------------------------------------'
Private GLODatSelected As Date 'A public property to get the user's selected date

Private GLOBooInstance As Boolean
Private GLOBooJustForm As Boolean
Private GLOBooMsgBox As Boolean
Private GLOBooMarkedDay As Boolean

Private GLOLngMonths As Double
Private GLOLngYears As Double
Private GLOLngDays As Long
Private GLOLngUnDays As Long

Private GLOLngDaysColor As Long

Private GLOLngDayBorderStyle As Long 'Default values
Private GLOLngDayBorderColor As Long
Private GLOLngDayBackColor As Long
Private GLOLngDayForeColor As Long

Private GLOLngCustomDayBackColor As Long 'Custom values
Private GLOLngCustomDayForeColor As Long

Private GLOLngDaySunForecolor As Long

' VARIABLES CUSTOM ------------------------------------------------------------------'
Public Enum DPC_JustForm
  dpcAsObject
  dpcAsForm
End Enum

Public Enum DPC_Alignment
  dpcLeft
  dpcRight
  dpcCenter
End Enum

' CONSTANTS (GLOSTR_) ---------------------------------------------------------------'
Private Const GLOLNG_DAY_SIDE_COLOR As Long = 7755409

' CONSTANTS ERROR MESSAGES (GLOSTR_ERR_) --------------------------------------------'
Private Const CUSTOM_ERROR As Long = VBA.vbObjectError + 516



' PROPERTIES ------------------------------------------------------------------------'



' CONSTRUCTOR/DESTRUCTOR ------------------------------------------------------------'
Private Sub Class_Initialize()
  
  Let GLOLngDays = VBA.Day(VBA.Date)
  Let GLOLngMonths = VBA.Month(VBA.Date)
  Let GLOLngYears = VBA.Year(VBA.Date)

  Let GLOLngDaysColor = VBA.RGB(100, 100, 100)
  Let GLOLngDaySunForecolor = VBA.RGB(255, 100, 100)
  
  Let GLOLngCustomDayBackColor = VBA.RGB(33, 115, 70)
  Let GLOLngCustomDayForeColor = GLOLngCustomDayBackColor + GLOLNG_DAY_SIDE_COLOR

  Let GLODatSelected = VBA.Date()

End Sub

Private Sub Class_Terminate()
  Call DatePicker_core_reset
End Sub



' CONTROLS EVENTS -------------------------------------------------------------------'
Private Sub GLOOBJ_SpinHead1_Change()
 Call DatePicker_core_change_month_year(GLOOBJ_SpinHead1.Value)
End Sub

Private Sub GLOOBJ_SpinHead2_Change()
 Call DatePicker_core_change_month_year(GLOOBJ_SpinHead2.Value, True)
End Sub

Private Sub GLOOBJ_LblHead2_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblHead2, True)
End Sub

Private Sub GLOOBJ_LblHead2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Let GLOOBJ_LblHead2.ForeColor = GLOLngCustomDayBackColor
End Sub


Private Sub GLOOBJ_LblBody1_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody1)
End Sub

Private Sub GLOOBJ_LblBody1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody1)
End Sub

Private Sub GLOOBJ_LblBody2_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody2)
End Sub

Private Sub GLOOBJ_LblBody2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody2)
End Sub

Private Sub GLOOBJ_LblBody3_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody3)
End Sub

Private Sub GLOOBJ_LblBody3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody3)
End Sub

Private Sub GLOOBJ_LblBody4_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody4)
End Sub

Private Sub GLOOBJ_LblBody4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody4)
End Sub

Private Sub GLOOBJ_LblBody5_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody5)
End Sub

Private Sub GLOOBJ_LblBody5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody5)
End Sub

Private Sub GLOOBJ_LblBody6_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody6)
End Sub

Private Sub GLOOBJ_LblBody6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody6)
End Sub

Private Sub GLOOBJ_LblBody7_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody7)
End Sub

Private Sub GLOOBJ_LblBody7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody7)
End Sub

Private Sub GLOOBJ_LblBody8_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody8)
End Sub

Private Sub GLOOBJ_LblBody8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody8)
End Sub

Private Sub GLOOBJ_LblBody9_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody9)
End Sub

Private Sub GLOOBJ_LblBody9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody9)
End Sub

Private Sub GLOOBJ_LblBody10_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody10)
End Sub

Private Sub GLOOBJ_LblBody10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody10)
End Sub

Private Sub GLOOBJ_LblBody11_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody11)
End Sub

Private Sub GLOOBJ_LblBody11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody11)
End Sub

Private Sub GLOOBJ_LblBody12_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody12)
End Sub

Private Sub GLOOBJ_LblBody12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody12)
End Sub

Private Sub GLOOBJ_LblBody13_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody13)
End Sub

Private Sub GLOOBJ_LblBody13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody13)
End Sub

Private Sub GLOOBJ_LblBody14_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody14)
End Sub

Private Sub GLOOBJ_LblBody14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody14)
End Sub

Private Sub GLOOBJ_LblBody15_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody15)
End Sub

Private Sub GLOOBJ_LblBody15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody15)
End Sub

Private Sub GLOOBJ_LblBody16_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody16)
End Sub

Private Sub GLOOBJ_LblBody16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody16)
End Sub

Private Sub GLOOBJ_LblBody17_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody17)
End Sub

Private Sub GLOOBJ_LblBody17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody17)
End Sub

Private Sub GLOOBJ_LblBody18_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody18)
End Sub

Private Sub GLOOBJ_LblBody18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody18)
End Sub

Private Sub GLOOBJ_LblBody19_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody19)
End Sub

Private Sub GLOOBJ_LblBody19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody19)
End Sub

Private Sub GLOOBJ_LblBody20_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody20)
End Sub

Private Sub GLOOBJ_LblBody20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody20)
End Sub

Private Sub GLOOBJ_LblBody21_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody21)
End Sub

Private Sub GLOOBJ_LblBody21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody21)
End Sub

Private Sub GLOOBJ_LblBody22_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody22)
End Sub

Private Sub GLOOBJ_LblBody22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody22)
End Sub

Private Sub GLOOBJ_LblBody23_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody23)
End Sub

Private Sub GLOOBJ_LblBody23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody23)
End Sub

Private Sub GLOOBJ_LblBody24_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody24)
End Sub

Private Sub GLOOBJ_LblBody24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody24)
End Sub

Private Sub GLOOBJ_LblBody25_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody25)
End Sub

Private Sub GLOOBJ_LblBody25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody25)
End Sub

Private Sub GLOOBJ_LblBody26_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody26)
End Sub

Private Sub GLOOBJ_LblBody26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody26)
End Sub

Private Sub GLOOBJ_LblBody27_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody27)
End Sub

Private Sub GLOOBJ_LblBody27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody27)
End Sub

Private Sub GLOOBJ_LblBody28_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody28)
End Sub

Private Sub GLOOBJ_LblBody28_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody28)
End Sub

Private Sub GLOOBJ_LblBody29_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody29)
End Sub

Private Sub GLOOBJ_LblBody29_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody29)
End Sub

Private Sub GLOOBJ_LblBody30_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody30)
End Sub

Private Sub GLOOBJ_LblBody30_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody30)
End Sub

Private Sub GLOOBJ_LblBody31_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody31)
End Sub

Private Sub GLOOBJ_LblBody31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody31)
End Sub

Private Sub GLOOBJ_LblBody32_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody32)
End Sub

Private Sub GLOOBJ_LblBody32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody32)
End Sub

Private Sub GLOOBJ_LblBody33_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody33)
End Sub

Private Sub GLOOBJ_LblBody33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody33)
End Sub

Private Sub GLOOBJ_LblBody34_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody34)
End Sub

Private Sub GLOOBJ_LblBody34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody34)
End Sub

Private Sub GLOOBJ_LblBody35_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody35)
End Sub

Private Sub GLOOBJ_LblBody35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody35)
End Sub

Private Sub GLOOBJ_LblBody36_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody36)
End Sub

Private Sub GLOOBJ_LblBody36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody36)
End Sub

Private Sub GLOOBJ_LblBody37_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody37)
End Sub

Private Sub GLOOBJ_LblBody37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody37)
End Sub

Private Sub GLOOBJ_LblBody38_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody38)
End Sub

Private Sub GLOOBJ_LblBody38_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody38)
End Sub

Private Sub GLOOBJ_LblBody39_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody39)
End Sub

Private Sub GLOOBJ_LblBody39_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody39)
End Sub

Private Sub GLOOBJ_LblBody40_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody40)
End Sub

Private Sub GLOOBJ_LblBody40_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody40)
End Sub

Private Sub GLOOBJ_LblBody41_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody41)
End Sub

Private Sub GLOOBJ_LblBody41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody41)
End Sub

Private Sub GLOOBJ_LblBody42_Click()
 Call DatePicker_controls_event_select_date(GLOOBJ_LblBody42)
End Sub

Private Sub GLOOBJ_LblBody42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_controls_event_mouse_move(GLOOBJ_LblBody42)
End Sub



' METHODS PRIVATE -------------------------------------------------------------------'
' CORE CONTROLS METHODS LIST (PRIVATE) ----------------------------------------------'
Private Sub DatePicker_open( _
  ByVal DFRM As Object, _
  ByVal DEnuJustForm As DPC_JustForm, _
  ByVal DVbaFirstDay As VbDayOfWeek, _
  ByVal DControl As Object, _
  ByVal DEnuAlign As DPC_Alignment, _
  ByVal DxlBaseColor As XlRgbColor, _
  ByVal DBooMsgBox As Boolean)
'  Optional DxlBaseColor As Long = 4616993, _

  #If Debugging Then
    Call Debug_assert((DFRM Is Nothing Or DControl Is Nothing), "DatePicker")
  #End If
  
  ' If second call, DPC already exists, so, close it
  If GLOBooInstance Then Call DatePicker_close: GoTo EH
  
  ' Create and Add DPC to Form controls
  Call DatePicker_core_add(DFRM, DEnuJustForm, DVbaFirstDay)
  
  ' Show DPC
  Call DatePicker_core_use(DEnuJustForm, DControl, DEnuAlign, DxlBaseColor, DBooMsgBox)

EH:
  If Not VBA.Err.Number = 0 Then
    Debug.Print VBA.Err.Description
    On Error GoTo -1
    Call VBA.Err.Clear
  End If

End Sub

Private Sub DatePicker_close()
  
  Dim DPCParent As Object
  
  On Error GoTo EH
  If GLOOBJ_FraDatepicker2 Is Nothing Then GoTo EH
  Set DPCParent = GLOOBJ_FraDatepicker1.Parent
  Let GLOOBJ_FraDatepicker1.Visible = False
  Call DPCParent.Controls.Remove("DatepickerFrame1")
  Call DatePicker_core_reset
  
  #If Debugging Then
    Dim cC As Object
    For Each cC In DPCParent.Controls
      Debug.Print cC.Name
    Next cC
    Set cC = Nothing
  #End If
  
  Set DPCParent = Nothing
    
EH:
  Set DPCParent = Nothing
  If Not VBA.Err.Number = 0 Then
    Debug.Print VBA.Err.Description
    Call VBA.Err.Clear
  End If

End Sub

Private Sub DatePicker_controls_event_select_date(ByVal CtrlLbl As MSForms.Label, Optional BooToday As Boolean = False)
  
  On Error GoTo EH
  Let CtrlLbl.ForeColor = GLOLngDayForeColor
  Let GLOOBJ_FraDatepicker2.Parent.Visible = False
  Let GLODatSelected = IIf(BooToday, VBA.Date(), VBA.Format(CtrlLbl.Caption, "00") & "/" & VBA.Format(GLOLngMonths, "00") & "/" & GLOLngYears)
  If GLOBooMsgBox Then Call VBA.MsgBox("Fecha: " & IIf(BooToday, VBA.Date(), VBA.Format(CtrlLbl.Caption, "00") & "/" & VBA.Format(GLOLngMonths, "00") & "/" & GLOLngYears))
  If Not BooToday Then Call DatePicker_core_remark_selected_day(CtrlLbl)
  
  If Not GLOOBJ_CtlReceptor Is Nothing Then
    Select Case VBA.TypeName(GLOOBJ_CtlReceptor)
      Case "TextBox"
        GLOOBJ_CtlReceptor.Text = GLODatSelected
      Case "ComboBox", "Range"
        GLOOBJ_CtlReceptor.Value = GLODatSelected
      Case "Label"
        GLOOBJ_CtlReceptor.Caption = GLODatSelected
    End Select
  End If
  
  If GLOBooJustForm Then
    Call VBA.Unload(GLOOBJ_FraDatepicker1.Parent)
  Else
    Call DatePicker_close
  End If

EH:

End Sub

Private Sub DatePicker_controls_event_mouse_move(ByVal CtrlLbl As MSForms.Label)
  
  Dim Ctrl As MSForms.Label

  On Error GoTo EH
  For Each Ctrl In GLOOBJ_FraDatepicker2.Controls
    Let Ctrl.ForeColor = IIf((VBA.CLng(Ctrl.Tag) Mod 7) = 0, GLOLngDaySunForecolor, GLOLngCustomDayForeColor)
  Next Ctrl
  Let CtrlLbl.ForeColor = IIf((VBA.CLng(CtrlLbl.Tag) Mod 7) = 0, VBA.vbRed, VBA.vbBlack)

EH:

End Sub



' CORE METHODS LIST PRIVATE ---------------------------------------------------------'
''********************************** NO MODIFICAR SI NO SABE **********************************
''*************************************|||||||||||||||||||*************************************
''***************************************|||||||||||||||***************************************
''*****************************************|||||||||||*****************************************
''*******************************************|||||||*******************************************
' ANDRÉS ROJAS MONCADA'S CALENDAR CORE ALGORITHM
' Nota del autor -----------------------------------------------------------------------------

' Creado por Andrés Rojas Moncada - Autor del canal Excel Hecho Fácil en YouTube

' Versión 1.0 - 20 de julio de 2015

' URL del canal: www.youtube.com/jarmoncada01

' Si quieres usarlo, solo copia y pega el presente módulo en conjunto con el UserForm y listo.

' Para ver algunos ejemplos sobre el uso de este calendario, observa este video.

' Enlace: |||||||||||||||||||| https://www.youtube.com/watch?v=FkjsuN2zqSU ||||||||||||||||||||

' ! Muchas gracias y espero lo disfruten ! ---------------------------------------------------

Private Sub DatePicker_core_InicializaFormularioCalendario()
  
  On Error GoTo EH
  Let GLOOBJ_FraDatepicker1.Controls("DatepickerLblHead1").Caption = VBA.Space$(5) & VBA.StrConv(VBA.MonthName(GLOLngMonths), vbProperCase) & ", " & GLOLngYears
  With GLOOBJ_FraDatepicker2.Controls("GLOOBJ_LblBody" & GLOLngDays)
    Let GLOLngDayBorderStyle = .BorderStyle
    Let GLOLngDayBorderColor = .BorderColor
    Let GLOLngDayBackColor = .BackColor
    Let GLOLngDayForeColor = .ForeColor
  End With
  
  Let GLOOBJ_SpinHead1.Value = VBA.Month(VBA.Date())
  Let GLOOBJ_SpinHead2.Value = VBA.Year(VBA.Date())
  Call DatePicker_core_CargarLosDias(GLOLngYears, GLOLngMonths)
EH:

End Sub

Private Sub DatePicker_core_CargarLosDias(ByVal Anho As Long, ByVal Mes As Long)
  
  Dim FechaDelPrimerDia As Date
  Dim FechaDelUltimoDia As Date
  Dim DiaSemanaPrimerDia As Long
  Dim VariableControl As MSForms.Control
  Dim Contador As Long
  
  On Error GoTo EH
  Let FechaDelPrimerDia = VBA.DateSerial(Anho, Mes, 1)
  Let FechaDelUltimoDia = Application.WorksheetFunction.EoMonth(VBA.DateSerial(Anho, Mes, 1), 0)
  Let DiaSemanaPrimerDia = Application.WorksheetFunction.Weekday(FechaDelPrimerDia, 2)
  Let Contador = 1

  For Each VariableControl In GLOOBJ_FraDatepicker2.Controls
    Let VariableControl.Visible = (VariableControl.Tag >= DiaSemanaPrimerDia And Contador <= VBA.Day(FechaDelUltimoDia))
    If (Contador = VBA.Day(VBA.Date())) And Not GLOBooMarkedDay Then
      Let GLOLngUnDays = Contador + DiaSemanaPrimerDia - 1
      Let GLOBooMarkedDay = True
    End If
    If VariableControl.Tag >= DiaSemanaPrimerDia And Contador <= VBA.Day(FechaDelUltimoDia) Then
      Let VariableControl.Caption = Contador
      Let Contador = Contador + 1
    End If
  Next VariableControl
  
EH:

End Sub
''********************************** NO MODIFICAR SI NO SABE **********************************
''*************************************|||||||||||||||||||*************************************
''***************************************|||||||||||||||***************************************
''*****************************************|||||||||||*****************************************
''*******************************************|||||||*******************************************
' ANDRÃ‰S ROJAS MONCADA'S CALENDAR CORE ALGORITHM                                ----- o ------

Private Sub DatePicker_core_change_month_year(ByVal DLngValue As Long, Optional DBooMonthYear As Boolean = False)
  
  On Error GoTo EH
  With GLOOBJ_SpinHead1
    If Not (VBA.IsNull(.Value)) Then
      If Not DBooMonthYear Then 'Monts
        With GLOOBJ_SpinHead1
          If DLngValue > 12 Then
            Let .Value = 1: GLOLngYears = GLOLngYears + 1
            If GLOLngYears >= 9999 Then GLOLngYears = 9999
          End If
          If DLngValue < 1 Then
            Let .Value = 12: GLOLngYears = GLOLngYears - 1
            If GLOLngYears <= 1900 Then GLOLngYears = 1900
          End If
          Let GLOLngMonths = .Value
          Let .Parent.Controls("DatepickerLblHead1").Caption = VBA.Space$(5) & VBA.StrConv(VBA.MonthName(GLOLngMonths), vbProperCase) & ", " & GLOLngYears
        End With
      Else 'Years
        With GLOOBJ_SpinHead2
          If DLngValue >= 9999 Then
            Let GLOLngYears = 9999
          ElseIf DLngValue <= 1900 Then
            Let GLOLngYears = 1900
          Else
            Let GLOLngYears = .Value
          End If
          Let .Parent.Controls("DatepickerLblHead1").Caption = VBA.Space$(5) & VBA.StrConv(VBA.MonthName(GLOLngMonths), vbProperCase) & ", " & GLOLngYears
        End With
      End If

      Call DatePicker_core_CargarLosDias(GLOLngYears, GLOLngMonths)
      Call DatePicker_core_remark_the_day
    End If
  End With

EH:

End Sub

Private Sub DatePicker_core_remark_the_day()
  
  On Error GoTo EH
  With GLOOBJ_FraDatepicker2.Controls("GLOOBJ_LblBody" & GLOLngUnDays)
    If GLOOBJ_SpinHead1.Value = VBA.Month(VBA.Date()) And GLOLngYears = VBA.Year(VBA.Date()) Then
      Let .BorderStyle = fmBorderStyleSingle
      Let .BorderColor = GLOLngDayBorderColor
      Let .BackColor = GLOLngCustomDayBackColor
      Let .ForeColor = GLOLngCustomDayForeColor
      Let .BackStyle = fmBackStyleOpaque
    Else
      Let .BorderStyle = GLOLngDayBorderStyle
      Let .BorderColor = GLOLngDayBorderColor
      Let .BackColor = GLOLngDayBackColor
      Let .ForeColor = GLOLngDayForeColor
      Let .BackStyle = fmBackStyleTransparent
    End If
  End With

EH:

End Sub

Private Sub DatePicker_core_remark_selected_day(ByVal DCtrlLblDay As MSForms.Label)
  
  Dim Ctrl As MSForms.Label

  On Error GoTo EH
  For Each Ctrl In GLOOBJ_FraDatepicker2.Controls
    With Ctrl
      Let .ForeColor = GLOLngDayForeColor
      Let .BackStyle = fmBackStyleTransparent
    End With
  Next Ctrl
  With GLOOBJ_FraDatepicker2.Controls("GLOOBJ_LblBody" & GLOLngUnDays)
    Let .ForeColor = GLOLngDayForeColor
    Let .BackStyle = fmBackStyleTransparent
  End With
  With DCtrlLblDay
    Let .BackColor = GLOLngCustomDayBackColor
    Let .ForeColor = GLOLngCustomDayForeColor ' VBA.vbWhite
    Let .BackStyle = fmBackStyleOpaque
  End With

EH:

End Sub

Private Sub DatePicker_core_change_colors()
  
  On Error GoTo EH
  With GLOOBJ_FraDatepicker1
    Let .Controls("DatepickerLblHead1").ForeColor = GLOLngCustomDayForeColor
    Let .Controls("DatepickerLblHead1").BackColor = GLOLngCustomDayBackColor
  End With

EH:

End Sub

Private Sub DatePicker_core_add(ByVal DFRM As Object, ByVal DEnuJustForm As DPC_JustForm, ByVal DVbaFirstDay As VbDayOfWeek)
  
  On Error GoTo EH
  Call DatePicker_core_set_controls(DFRM, DEnuJustForm, DVbaFirstDay)

EH:

End Sub

Private Sub DatePicker_core_use( _
  ByVal DEnuJustForm As DPC_JustForm, _
  ByVal DControl As Object, _
  ByVal DEnuAlign As DPC_Alignment, _
  ByVal DxlBaseColor As XlRgbColor, _
  ByVal DBooMsgBox As Boolean)

'  Optional DxlBaseColor As Long = 4616993

  Dim LngHeight As Long, LngWidth As Long
  'Dim PosCoord As positinons
  
  On Error GoTo EH
  Let GLOLngDayBorderColor = DxlBaseColor
  Let GLOLngCustomDayBackColor = DxlBaseColor
  Let GLOLngCustomDayForeColor = DxlBaseColor + GLOLNG_DAY_SIDE_COLOR ' Triade, soft color

  Let GLOBooJustForm = DEnuJustForm
  Let GLOBooMsgBox = DBooMsgBox
  
  On Error Resume Next
  With GLOOBJ_FraDatepicker1.Parent ' TheForm or object parent
    
    Set GLOOBJ_CtlReceptor = DControl
    If DEnuJustForm Then
      
      Let .Controls("DatepickerFrame1").Left = 0
      Let .Controls("DatepickerFrame1").Top = 0
      Let .Controls("DatepickerFrame1").Visible = True ' Not .Controls("DatepickerFrame1").Visible
      Let LngWidth = .Controls("DatepickerFrame1").Width
      Let LngHeight = .Controls("DatepickerFrame1").Height
      Let .Width = LngWidth * 1.09
      Let .Height = LngHeight * 1.2
      If VBA.TypeName(DControl) = "Range" Then
        ' This is a complex task, because coordinates of Excel sheets are very variables _
          but is possible with some algorithms like Form Positioner by Pearson Software Consulting Services. _
          More information: http://www.cpearson.com/excel/FormPosition.htm
        ' Here will apply some logical operations to aproximatelly obtain the cell left and top to our DPC as UserForm.
        Let .StartUpPosition = 0
        Let .Left = (DControl.Left + DControl.Width) * 1.13
        Let .Top = (.Height * 0.96) + DControl.Top + DControl.Height
      End If
      Call DatePicker_core_change_colors
      Let .Caption = "DatePickerClass"
      Call VBA.Load(GLOOBJ_FraDatepicker1.Parent)
      Call .Show
    
    Else
      
      If Not DControl Is Nothing Then
        Let .Controls("DatepickerFrame1").Left = Datepicker_asset_get_horizontal_align(DControl, DEnuAlign)
        Let .Controls("DatepickerFrame1").Top = Datepicker_asset_get_vertical_align(DControl)
        Let .Controls("DatepickerFrame1").Visible = True
        Call DatePicker_core_change_colors
      End If
    
    End If
    If Not GLOOBJ_FraDatepicker2 Is Nothing Then Let GLOOBJ_FraDatepicker1.Visible = True
  
  End With

EH:

End Sub

Private Sub DatePicker_core_set_controls(ByVal DFRM As Object, ByVal DEnuJustForm As DPC_JustForm, ByVal DVbaFirstDay As VbDayOfWeek)
  
  'BUILDS A FRAME WITH ALL CONTROLS FOR CALENDAR
  ' [x]  A Header Label for Month, Year.
  ' [x]  A SpinButton to navigate across Months/Years.
  ' [x]  A SubHeader Label for Today: dd/mm/yy.
  ' [x]  A SubHeader Label for Each first letter of the day.
  ' [x]  A Frame for all Days.
  ' [x]  And all labels for all Days.
  Dim ObjFrm As Object
  
  Dim BooIsFrm As Boolean
  Dim i As Integer
  Dim j As Single, xCol As Single, xRow As Single
  
  ' Verifies required Form
  If DFRM Is Nothing Then Call VBA.Err.Raise(CUSTOM_ERROR, "DPC::-set_controls", "¡A UserForm object is required!")
  
  
  'Verifies if DFRM is a correct object
  Let BooIsFrm = False
  If VBA.UserForms.Count > 0 Then
    For Each ObjFrm In VBA.UserForms
      If VBA.TypeName(DFRM) = VBA.TypeName(ObjFrm) Then Let BooIsFrm = True: Exit For
    Next ObjFrm
    Set ObjFrm = Nothing
  End If
  ' It is required that DPC parent object would be a UserForm
  If Not BooIsFrm Then Call VBA.Err.Raise(CUSTOM_ERROR, "DPC::-set_controls", "¡A UserForm object is required!")
  ' If DPC is required as a Form in a Excel's workbook, its object parent must be always a UserForm.
  If DEnuJustForm = dpcAsForm And Not BooIsFrm Then Call VBA.Err.Raise(CUSTOM_ERROR, "DPC::-set_controls", "¡A UserForm object is required!")
  
  
  ' Lets create the DPC
  On Error GoTo EH
  With DFRM
    Set GLOOBJ_FraDatepicker1 = .Controls.Add("Forms.Frame.1", "DatepickerFrame1", False)
    With GLOOBJ_FraDatepicker1 '.Controls("DatepickerFrame1")
      Let .Width = 114
      Let .Height = 130
      Let .BorderStyle = fmBorderStyleSingle
      
      Call .Controls.Add("Forms.Label.1", "DatepickerLblHead1")
      Set GLOOBJ_LblHead2 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblHead2")
      
      Set GLOOBJ_SpinHead1 = .Controls.Add("Forms.SpinButton.1", "GLOOBJ_SpinHead1")
      Set GLOOBJ_SpinHead2 = .Controls.Add("Forms.SpinButton.1", "GLOOBJ_SpinHead2")
      
      Set GLOOBJ_FraDatepicker2 = .Controls.Add("Forms.Frame.1", "GLOOBJ_FraDatepicker2")
      
      ' Place the DPC layout
      With GLOOBJ_SpinHead1
        Let .Height = 12 * 0.5
        Let .Width = .Parent.Width * 0.2
        Let .Left = .Parent.Width - (.Width + 10)
        Let .Delay = 50
        Let .MIN = 0
        Let .Max = 13
      End With
      With GLOOBJ_SpinHead2
        Let .Height = 12 * 0.5
        Let .Width = .Parent.Width * 0.2
        Let .Left = GLOOBJ_SpinHead1.Left
        Let .Top = GLOOBJ_SpinHead1.Top + GLOOBJ_SpinHead1.Height
        Let .Delay = 50
        Let .MIN = 0 '1900
        Let .Max = 9999
      End With
    
      With .Controls("DatepickerLblHead1")
        Let .Height = 12
        Let .Width = .Parent.Width
        Let .Caption = "     Enero, 2020"
        Let .Font.Bold = True
      End With
      With .Controls("GLOOBJ_LblHead2")
        Let .Caption = VBA.Space$(5) & "Hoy: " & VBA.Date
        Let .Top = 12
        Let .BackStyle = fmBackStyleTransparent
        Let .ForeColor = GLOLngDaysColor ' VBA.RGB(100, 100, 100)
      End With
      
      For i = 0 To 6
        Call .Controls.Add("Forms.Label.1", "DatepickerLblHead3" & i)
        With .Controls("DatepickerLblHead3" & i)
          Let .Caption = VBA.UCase(VBA.Left(VBA.WeekdayName(VBA.Weekday(i + DVbaFirstDay, DVbaFirstDay)), 1)) '"      L   M    M   J    V    S   D"
          Let .TextAlign = fmTextAlignCenter
          Let .Font.Bold = True
          If VBA.Weekday(i + DVbaFirstDay, DVbaFirstDay) = VBA.vbSaturday Then .ForeColor = VBA.vbRed
          Let .BackStyle = fmBackStyleTransparent
          Let .Height = 9 '10.5
          Let .Width = 12 'FrameCal1.Width + 4
          Let .Left = 9 + (i * .Width * 1.15) '-2
          Let .Top = 24
        End With
      Next i
      
      Call .ZOrder(0)
    End With
    
    With GLOOBJ_FraDatepicker2
      Set GLOOBJ_LblBody1 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody1")
      Set GLOOBJ_LblBody2 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody2")
      Set GLOOBJ_LblBody3 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody3")
      Set GLOOBJ_LblBody4 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody4")
      Set GLOOBJ_LblBody5 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody5")
      Set GLOOBJ_LblBody6 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody6")
      Set GLOOBJ_LblBody7 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody7")

      Set GLOOBJ_LblBody8 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody8")
      Set GLOOBJ_LblBody9 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody9")
      Set GLOOBJ_LblBody10 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody10")
      Set GLOOBJ_LblBody11 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody11")
      Set GLOOBJ_LblBody12 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody12")
      Set GLOOBJ_LblBody13 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody13")
      Set GLOOBJ_LblBody14 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody14")

      Set GLOOBJ_LblBody15 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody15")
      Set GLOOBJ_LblBody16 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody16")
      Set GLOOBJ_LblBody17 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody17")
      Set GLOOBJ_LblBody18 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody18")
      Set GLOOBJ_LblBody19 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody19")
      Set GLOOBJ_LblBody20 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody20")
      Set GLOOBJ_LblBody21 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody21")

      Set GLOOBJ_LblBody22 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody22")
      Set GLOOBJ_LblBody23 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody23")
      Set GLOOBJ_LblBody24 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody24")
      Set GLOOBJ_LblBody25 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody25")
      Set GLOOBJ_LblBody26 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody26")
      Set GLOOBJ_LblBody27 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody27")
      Set GLOOBJ_LblBody28 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody28")

      Set GLOOBJ_LblBody29 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody29")
      Set GLOOBJ_LblBody30 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody30")
      Set GLOOBJ_LblBody31 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody31")
      Set GLOOBJ_LblBody32 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody32")
      Set GLOOBJ_LblBody33 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody33")
      Set GLOOBJ_LblBody34 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody34")
      Set GLOOBJ_LblBody35 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody35")

      Set GLOOBJ_LblBody36 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody36")
      Set GLOOBJ_LblBody37 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody37")
      Set GLOOBJ_LblBody38 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody38")
      Set GLOOBJ_LblBody39 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody39")
      Set GLOOBJ_LblBody40 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody40")
      Set GLOOBJ_LblBody41 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody41")
      Set GLOOBJ_LblBody42 = .Controls.Add("Forms.Label.1", "GLOOBJ_LblBody42")
      
      Let .Top = 35
      Let .Left = -1
      Let xCol = 9 '10
      Let xRow = 5 '40
      Let j = 1
      For i = 1 To 42
        With .Controls("GLOOBJ_LblBody" & i)
          Let .Caption = i
          Let .Width = 12
          Let .Height = 12
          Let .TextAlign = fmTextAlignCenter
          Let .Top = xRow '10
          Let .Left = xCol '(i * 12) + 5 ' (.Width + 4)
          Let .BackStyle = fmBackStyleTransparent
          Let .ForeColor = GLOLngCustomDayForeColor ' VBA.RGB(100, 100, 100)
          Let .Tag = i
          
          If (i Mod 7) = 0 Then
            Let xCol = 9
            Let xRow = xRow + .Height * 1.15
            Let j = 1
            Let .ForeColor = VBA.RGB(255, 100, 100) ' VBA.vbRed
          Else
            Let xCol = 9 + (j * .Width * 1.15) 'col + 1
            Let j = j + 1
          End If
        End With
      Next i
    End With
  End With
  
  Call DatePicker_core_InicializaFormularioCalendario
  Let GLOBooInstance = True

EH:
  Set ObjFrm = Nothing

End Sub

Private Sub DatePicker_core_reset()
  
  ' Destoys all objects setted in DPC
  On Error GoTo EH
  Set GLOOBJ_LblHead2 = Nothing

  Set GLOOBJ_SpinHead1 = Nothing
  Set GLOOBJ_SpinHead2 = Nothing

  Set GLOOBJ_LblBody1 = Nothing
  Set GLOOBJ_LblBody2 = Nothing
  Set GLOOBJ_LblBody3 = Nothing
  Set GLOOBJ_LblBody4 = Nothing
  Set GLOOBJ_LblBody5 = Nothing
  Set GLOOBJ_LblBody6 = Nothing
  Set GLOOBJ_LblBody7 = Nothing

  Set GLOOBJ_LblBody8 = Nothing
  Set GLOOBJ_LblBody9 = Nothing
  Set GLOOBJ_LblBody10 = Nothing
  Set GLOOBJ_LblBody11 = Nothing
  Set GLOOBJ_LblBody12 = Nothing
  Set GLOOBJ_LblBody13 = Nothing
  Set GLOOBJ_LblBody14 = Nothing

  Set GLOOBJ_LblBody15 = Nothing
  Set GLOOBJ_LblBody16 = Nothing
  Set GLOOBJ_LblBody17 = Nothing
  Set GLOOBJ_LblBody18 = Nothing
  Set GLOOBJ_LblBody19 = Nothing
  Set GLOOBJ_LblBody20 = Nothing
  Set GLOOBJ_LblBody21 = Nothing

  Set GLOOBJ_LblBody22 = Nothing
  Set GLOOBJ_LblBody23 = Nothing
  Set GLOOBJ_LblBody24 = Nothing
  Set GLOOBJ_LblBody25 = Nothing
  Set GLOOBJ_LblBody26 = Nothing
  Set GLOOBJ_LblBody27 = Nothing
  Set GLOOBJ_LblBody28 = Nothing

  Set GLOOBJ_LblBody29 = Nothing
  Set GLOOBJ_LblBody30 = Nothing
  Set GLOOBJ_LblBody31 = Nothing
  Set GLOOBJ_LblBody32 = Nothing
  Set GLOOBJ_LblBody33 = Nothing
  Set GLOOBJ_LblBody34 = Nothing
  Set GLOOBJ_LblBody35 = Nothing

  Set GLOOBJ_LblBody36 = Nothing
  Set GLOOBJ_LblBody37 = Nothing
  Set GLOOBJ_LblBody38 = Nothing
  Set GLOOBJ_LblBody39 = Nothing
  Set GLOOBJ_LblBody40 = Nothing
  Set GLOOBJ_LblBody41 = Nothing
  Set GLOOBJ_LblBody42 = Nothing
  
  Set GLOOBJ_FraDatepicker2 = Nothing
  Set GLOOBJ_FraDatepicker1 = Nothing
  
  Let GLOBooInstance = False
EH:
End Sub



' ASSETS ----------------------------------------------------------------------------'
Private Function Datepicker_asset_get_horizontal_align(ByVal DCtlTarget As MSForms.Control, ByVal DEnuAlign As DPC_Alignment) As Long
  
  Dim xLeft As Long, xDPWidth As Long, CtlTargetParentLeft As Long
  
  On Error GoTo EH
  With GLOOBJ_FraDatepicker1.Parent
    
    Let CtlTargetParentLeft = IIf(VBA.TypeName(DCtlTarget.Parent) = VBA.TypeName(GLOOBJ_FraDatepicker1.Parent), 0, DCtlTarget.Parent.Left) + 2
    
    Let xDPWidth = GLOOBJ_FraDatepicker1.Width
    Select Case DEnuAlign
      Case dpcRight: Let xLeft = ((DCtlTarget.Left + DCtlTarget.Width) - xDPWidth) + CtlTargetParentLeft
      Case dpcCenter: Let xLeft = ((DCtlTarget.Left + (DCtlTarget.Width / 2)) - (xDPWidth / 2)) + (CtlTargetParentLeft / 2)
      Case Else: Let xLeft = DCtlTarget.Left + CtlTargetParentLeft  ' dpcLeft
    End Select
    
    If (xLeft + xDPWidth) > .InsideWidth Then Let xLeft = xLeft - ((xLeft + xDPWidth) - .InsideWidth)
    If xLeft < 0 Then Let xLeft = 0
  
  End With
  Let Datepicker_asset_get_horizontal_align = xLeft

EH:

End Function

Private Function Datepicker_asset_get_vertical_align(ByVal DCtlTarget As MSForms.Control) As Long
  
  Dim xTop As Long, xDPHeight As Long, CtlTargetParentTop As Long
  
  On Error GoTo EH
  With GLOOBJ_FraDatepicker1.Parent
    Let xDPHeight = GLOOBJ_FraDatepicker1.Height ' .Controls("DatepickerFrame1").Height
    
    Let CtlTargetParentTop = IIf(VBA.TypeName(DCtlTarget.Parent) = VBA.TypeName(GLOOBJ_FraDatepicker1.Parent), 0, DCtlTarget.Parent.Top) + 5
    
    Let xTop = DCtlTarget.Top + DCtlTarget.Height + CtlTargetParentTop
    
    If (xTop + xDPHeight) > .InsideHeight Then Let xTop = xTop - (((xTop + xDPHeight) - .InsideHeight) - DCtlTarget.Height)
  End With
  Let Datepicker_asset_get_vertical_align = xTop

EH:

End Function



' ERRORS MANNAGEMENT AND ASSERTSIONS ------------------------------------------------'
Private Sub Debug_assert(ByVal Cond As Boolean, Optional Origin As String = VBA.vbNullString)
  
  'This method works with a project variable called Debugging, _
   this variable can be initialized in menu: Tools > Project properties > Conditional compilation arguments: Debugging = 1
  #If Debbuging Then
    Debug.Assert Cond
    If Not Cond Then Debug.Print "Assert on: " & Origin
  #End If

End Sub



' PUBLIC METHODS LIST (INTERFACE) ---------------------------------------------------'
Public Sub DatePicker( _
  ByVal FRM As MSForms.UserForm, _
  Optional DJustForm As DPC_JustForm = dpcAsObject, _
  Optional DFirstDay As VbDayOfWeek = VBA.vbMonday, _
  Optional DControl As Object, _
  Optional DAlign As DPC_Alignment = dpcLeft, _
  Optional DBaseColor As XlRgbColor = 4616993, _
  Optional DMsgBox As Boolean = False)

  On Error GoTo EH
  Call DatePicker_open(FRM, DJustForm, DFirstDay, DControl, DAlign, DBaseColor, DMsgBox)

EH:

End Sub
' PUBLIC METHODS LIST (INTERFACE) ---------------------------------------------------'
' =========================== DATE PICKER CLASS STRUCTURE ============================'

