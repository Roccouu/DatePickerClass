VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatePickerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ------------------------------------------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2020 Roberto Carlos Romay Medina (Roccou)
' DatePickerClass, Excel simple date picker for use with macros.
' Version: 2.0.0
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
'
'
' ACKNOWLEDGMENS
' INSPIRED ON ANDRÉS ROJAS MONCADA'S CALENDAR CORE ALGORITHM - Thanks a lot AndrÃ©s for all your great talent.
' Nota del autor -----------------------------------------------------------------------------
' Andrés Rojas Moncada - Autor del canal Excel Hecho Fácil en YouTube
' Versión 1.0 - 20 de julio de 2015
' URL del canal: www.youtube.com/jarmoncada01
' Para ver algunos ejemplos sobre el uso de este calendario, observa este video.
' Enlace: |||||||||||||||||||| https://www.youtube.com/watch?v=FkjsuN2zqSU ||||||||||||||||||||
' ! Muchas gracias y espero lo disfruten ! ---------------------------------------------------
' ------------------------------------------------------------------------------------------------------------

Private DatepickerFrame2 As MSForms.Frame
Attribute DatepickerFrame2.VB_VarHelpID = -1

Private WithEvents DatepickerLblHead2 As MSForms.Label
Attribute DatepickerLblHead2.VB_VarHelpID = -1

Private WithEvents DatepickerSpinHead1 As MSForms.SpinButton
Attribute DatepickerSpinHead1.VB_VarHelpID = -1
Private WithEvents DatepickerSpinHead2 As MSForms.SpinButton
Attribute DatepickerSpinHead2.VB_VarHelpID = -1

Private WithEvents DatepickerLblBody1 As MSForms.Label
Attribute DatepickerLblBody1.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody2 As MSForms.Label
Attribute DatepickerLblBody2.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody3 As MSForms.Label
Attribute DatepickerLblBody3.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody4 As MSForms.Label
Attribute DatepickerLblBody4.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody5 As MSForms.Label
Attribute DatepickerLblBody5.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody6 As MSForms.Label
Attribute DatepickerLblBody6.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody7 As MSForms.Label
Attribute DatepickerLblBody7.VB_VarHelpID = -1

Private WithEvents DatepickerLblBody8 As MSForms.Label
Attribute DatepickerLblBody8.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody9 As MSForms.Label
Attribute DatepickerLblBody9.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody10 As MSForms.Label
Attribute DatepickerLblBody10.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody11 As MSForms.Label
Attribute DatepickerLblBody11.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody12 As MSForms.Label
Attribute DatepickerLblBody12.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody13 As MSForms.Label
Attribute DatepickerLblBody13.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody14 As MSForms.Label
Attribute DatepickerLblBody14.VB_VarHelpID = -1

Private WithEvents DatepickerLblBody15 As MSForms.Label
Attribute DatepickerLblBody15.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody16 As MSForms.Label
Attribute DatepickerLblBody16.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody17 As MSForms.Label
Attribute DatepickerLblBody17.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody18 As MSForms.Label
Attribute DatepickerLblBody18.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody19 As MSForms.Label
Attribute DatepickerLblBody19.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody20 As MSForms.Label
Attribute DatepickerLblBody20.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody21 As MSForms.Label
Attribute DatepickerLblBody21.VB_VarHelpID = -1

Private WithEvents DatepickerLblBody22 As MSForms.Label
Attribute DatepickerLblBody22.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody23 As MSForms.Label
Attribute DatepickerLblBody23.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody24 As MSForms.Label
Attribute DatepickerLblBody24.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody25 As MSForms.Label
Attribute DatepickerLblBody25.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody26 As MSForms.Label
Attribute DatepickerLblBody26.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody27 As MSForms.Label
Attribute DatepickerLblBody27.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody28 As MSForms.Label
Attribute DatepickerLblBody28.VB_VarHelpID = -1

Private WithEvents DatepickerLblBody29 As MSForms.Label
Attribute DatepickerLblBody29.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody30 As MSForms.Label
Attribute DatepickerLblBody30.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody31 As MSForms.Label
Attribute DatepickerLblBody31.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody32 As MSForms.Label
Attribute DatepickerLblBody32.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody33 As MSForms.Label
Attribute DatepickerLblBody33.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody34 As MSForms.Label
Attribute DatepickerLblBody34.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody35 As MSForms.Label
Attribute DatepickerLblBody35.VB_VarHelpID = -1

Private WithEvents DatepickerLblBody36 As MSForms.Label
Attribute DatepickerLblBody36.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody37 As MSForms.Label
Attribute DatepickerLblBody37.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody38 As MSForms.Label
Attribute DatepickerLblBody38.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody39 As MSForms.Label
Attribute DatepickerLblBody39.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody40 As MSForms.Label
Attribute DatepickerLblBody40.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody41 As MSForms.Label
Attribute DatepickerLblBody41.VB_VarHelpID = -1
Private WithEvents DatepickerLblBody42 As MSForms.Label
Attribute DatepickerLblBody42.VB_VarHelpID = -1


Private CDateSelected As Date 'A public property to get the user's selected date

Private cInstance As Boolean
Private cJustForm As Boolean
Private cMsgBox As Boolean
Private cReceptorControl As MSForms.Control

Private cMonths As Double
Private cYears As Double
Private cDays As Long
Private cUnDays As Long

Private cMarkedDay As Boolean

Private DaysColor As Long

Private DayBorderStyle As Long 'Default values
Private DayBorderColor As Long
Private DayBackColor As Long
Private DayForeColor As Long

Private cDayBackColor As Long 'Custom values
Private cDayForeColor As Long
Private Const cDaySideColor As Long = 7755409

Private DaySunForecolor As Long




'PROPERTIES
Public Property Get GetDATE() As Date
  GetDATE = CDateSelected
End Property



'CONSTRUCTOR/DESTRUCTOR
Private Sub Class_Initialize()
  cDays = VBA.Day(VBA.Date)
  cMonths = VBA.Month(VBA.Date)
  cYears = VBA.Year(VBA.Date)

  DaysColor = VBA.RGB(100, 100, 100)
  DaySunForecolor = VBA.RGB(255, 100, 100)
  
  cDayBackColor = VBA.RGB(33, 115, 70)
  cDayForeColor = cDayBackColor + cDaySideColor

  CDateSelected = VBA.Date()
End Sub

Private Sub Class_Terminate()
  Call DatePicker_reset
End Sub



'METHODS PRIVATE
'Methods Controls Events
'Methods calendar labels picker Events
Private Sub DatepickerSpinHead1_Change()
 Call DatePicker_ChangeMonthYear(DatepickerSpinHead1.Value)
End Sub

Private Sub DatepickerSpinHead2_Change()
 Call DatePicker_ChangeMonthYear(DatepickerSpinHead2.Value, True)
End Sub

Private Sub DatepickerLblHead2_Click()
 Call DatePicker_EventSelectDate(DatepickerLblHead2, True)
End Sub

Private Sub DatepickerLblHead2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 DatepickerLblHead2.ForeColor = cDayBackColor
End Sub


Private Sub DatepickerLblBody1_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody1)
End Sub

Private Sub DatepickerLblBody1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody1)
End Sub

Private Sub DatepickerLblBody2_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody2)
End Sub

Private Sub DatepickerLblBody2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody2)
End Sub

Private Sub DatepickerLblBody3_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody3)
End Sub

Private Sub DatepickerLblBody3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody3)
End Sub

Private Sub DatepickerLblBody4_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody4)
End Sub

Private Sub DatepickerLblBody4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody4)
End Sub

Private Sub DatepickerLblBody5_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody5)
End Sub

Private Sub DatepickerLblBody5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody5)
End Sub

Private Sub DatepickerLblBody6_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody6)
End Sub

Private Sub DatepickerLblBody6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody6)
End Sub

Private Sub DatepickerLblBody7_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody7)
End Sub

Private Sub DatepickerLblBody7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody7)
End Sub

Private Sub DatepickerLblBody8_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody8)
End Sub

Private Sub DatepickerLblBody8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody8)
End Sub

Private Sub DatepickerLblBody9_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody9)
End Sub

Private Sub DatepickerLblBody9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody9)
End Sub

Private Sub DatepickerLblBody10_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody10)
End Sub

Private Sub DatepickerLblBody10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody10)
End Sub

Private Sub DatepickerLblBody11_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody11)
End Sub

Private Sub DatepickerLblBody11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody11)
End Sub

Private Sub DatepickerLblBody12_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody12)
End Sub

Private Sub DatepickerLblBody12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody12)
End Sub

Private Sub DatepickerLblBody13_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody13)
End Sub

Private Sub DatepickerLblBody13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody13)
End Sub

Private Sub DatepickerLblBody14_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody14)
End Sub

Private Sub DatepickerLblBody14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody14)
End Sub

Private Sub DatepickerLblBody15_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody15)
End Sub

Private Sub DatepickerLblBody15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody15)
End Sub

Private Sub DatepickerLblBody16_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody16)
End Sub

Private Sub DatepickerLblBody16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody16)
End Sub

Private Sub DatepickerLblBody17_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody17)
End Sub

Private Sub DatepickerLblBody17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody17)
End Sub

Private Sub DatepickerLblBody18_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody18)
End Sub

Private Sub DatepickerLblBody18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody18)
End Sub

Private Sub DatepickerLblBody19_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody19)
End Sub

Private Sub DatepickerLblBody19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody19)
End Sub

Private Sub DatepickerLblBody20_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody20)
End Sub

Private Sub DatepickerLblBody20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody20)
End Sub

Private Sub DatepickerLblBody21_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody21)
End Sub

Private Sub DatepickerLblBody21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody21)
End Sub

Private Sub DatepickerLblBody22_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody22)
End Sub

Private Sub DatepickerLblBody22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody22)
End Sub

Private Sub DatepickerLblBody23_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody23)
End Sub

Private Sub DatepickerLblBody23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody23)
End Sub

Private Sub DatepickerLblBody24_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody24)
End Sub

Private Sub DatepickerLblBody24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody24)
End Sub

Private Sub DatepickerLblBody25_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody25)
End Sub

Private Sub DatepickerLblBody25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody25)
End Sub

Private Sub DatepickerLblBody26_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody26)
End Sub

Private Sub DatepickerLblBody26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody26)
End Sub

Private Sub DatepickerLblBody27_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody27)
End Sub

Private Sub DatepickerLblBody27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody27)
End Sub

Private Sub DatepickerLblBody28_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody28)
End Sub

Private Sub DatepickerLblBody28_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody28)
End Sub

Private Sub DatepickerLblBody29_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody29)
End Sub

Private Sub DatepickerLblBody29_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody29)
End Sub

Private Sub DatepickerLblBody30_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody30)
End Sub

Private Sub DatepickerLblBody30_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody30)
End Sub

Private Sub DatepickerLblBody31_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody31)
End Sub

Private Sub DatepickerLblBody31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody31)
End Sub

Private Sub DatepickerLblBody32_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody32)
End Sub

Private Sub DatepickerLblBody32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody32)
End Sub

Private Sub DatepickerLblBody33_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody33)
End Sub

Private Sub DatepickerLblBody33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody33)
End Sub

Private Sub DatepickerLblBody34_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody34)
End Sub

Private Sub DatepickerLblBody34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody34)
End Sub

Private Sub DatepickerLblBody35_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody35)
End Sub

Private Sub DatepickerLblBody35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody35)
End Sub

Private Sub DatepickerLblBody36_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody36)
End Sub

Private Sub DatepickerLblBody36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody36)
End Sub

Private Sub DatepickerLblBody37_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody37)
End Sub

Private Sub DatepickerLblBody37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody37)
End Sub

Private Sub DatepickerLblBody38_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody38)
End Sub

Private Sub DatepickerLblBody38_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody38)
End Sub

Private Sub DatepickerLblBody39_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody39)
End Sub

Private Sub DatepickerLblBody39_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody39)
End Sub

Private Sub DatepickerLblBody40_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody40)
End Sub

Private Sub DatepickerLblBody40_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody40)
End Sub

Private Sub DatepickerLblBody41_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody41)
End Sub

Private Sub DatepickerLblBody41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody41)
End Sub

Private Sub DatepickerLblBody42_Click()
 Call DatePicker_EventSelectDate(DatepickerLblBody42)
End Sub

Private Sub DatepickerLblBody42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 Call DatePicker_EventMouseMove(DatepickerLblBody42)
End Sub

'Methods Core Class
Private Sub DatePicker_SetControls(ByVal Frm As Object, Optional cFirstDay As Long = VBA.vbMonday)
  'BUILDS A FRAME WITH ALL CONTROLS FOR CALENDAR
  ' [x]  A Header Label for Month, Year.
  ' [x]  A SpinButton to navigate across Months/Years.
  ' [x]  A SubHeader Label for Today: dd/mm/yy.
  ' [x]  A SubHeader Label for Each first letter of the day.
  ' [x]  A Frame for all Days.
  ' [x]  And all labels for all Days.
  
  Dim i As Integer
  Dim j As Single
  Dim col As Single
  Dim row As Single
  
  If Frm Is Nothing Then GoTo DP_SC
  
  On Error GoTo DP_SC
  With Frm
    Call .Controls.Add("Forms.Frame.1", "DatepickerFrame1", False)
    With .Controls("DatepickerFrame1") 'FrameCal1
      .Width = 114
      .Height = 130
      .BorderStyle = fmBorderStyleSingle
    
      Call .Controls.Add("Forms.Label.1", "DatepickerLblHead1")
      Set DatepickerLblHead2 = .Controls.Add("Forms.Label.1", "DatepickerLblHead2")
      
      Set DatepickerSpinHead1 = .Controls.Add("Forms.SpinButton.1", "DatepickerSpinHead1")
      Set DatepickerSpinHead2 = .Controls.Add("Forms.SpinButton.1", "DatepickerSpinHead2")
    
      Set DatepickerFrame2 = .Controls.Add("Forms.Frame.1", "DatepickerFrame2")
      
      With DatepickerSpinHead1
        .Height = 12 * 0.5
        .Width = .Parent.Width * 0.2
        .Left = .Parent.Width - (.Width + 10)
        .Delay = 50
        .Min = 0
        .Max = 13
      End With
      With DatepickerSpinHead2
        .Height = 12 * 0.5
        .Width = .Parent.Width * 0.2
        .Left = DatepickerSpinHead1.Left
        .Top = DatepickerSpinHead1.Top + DatepickerSpinHead1.Height
        .Delay = 50
        .Min = 0 '1900
        .Max = 9999
      End With
    
      With .Controls("DatepickerLblHead1")
        .Height = 12
        .Width = .Parent.Width
        .Caption = "     Enero, 2020"
        .Font.Bold = True
      End With
      With .Controls("DatepickerLblHead2")
        .Caption = "     Hoy: " & VBA.Date
        .Top = 12
        .BackStyle = fmBackStyleTransparent
        .ForeColor = DaysColor ' VBA.RGB(100, 100, 100)
      End With
    
      For i = 0 To 6
        Call .Controls.Add("Forms.Label.1", "DatepickerLblHead3" & i)
        With .Controls("DatepickerLblHead3" & i)
          .Caption = VBA.UCase(VBA.Left(VBA.WeekdayName(VBA.Weekday(i + cFirstDay, cFirstDay)), 1)) '"      L   M    M   J    V    S   D"
          .TextAlign = fmTextAlignCenter
          .Font.Bold = True
          If VBA.Weekday(i + cFirstDay, cFirstDay) = VBA.vbSaturday Then .ForeColor = VBA.vbRed
          .BackStyle = fmBackStyleTransparent
          .Height = 9 '10.5
          .Width = 12 'FrameCal1.Width + 4
          .Left = 9 + (i * .Width * 1.15) '-2
          .Top = 24
        End With
      Next i
      
      Call .ZOrder(0)
    End With
    
    With DatepickerFrame2
      Set DatepickerLblBody1 = .Controls.Add("Forms.Label.1", "DatepickerLblBody1")
      Set DatepickerLblBody2 = .Controls.Add("Forms.Label.1", "DatepickerLblBody2")
      Set DatepickerLblBody3 = .Controls.Add("Forms.Label.1", "DatepickerLblBody3")
      Set DatepickerLblBody4 = .Controls.Add("Forms.Label.1", "DatepickerLblBody4")
      Set DatepickerLblBody5 = .Controls.Add("Forms.Label.1", "DatepickerLblBody5")
      Set DatepickerLblBody6 = .Controls.Add("Forms.Label.1", "DatepickerLblBody6")
      Set DatepickerLblBody7 = .Controls.Add("Forms.Label.1", "DatepickerLblBody7")
      
      Set DatepickerLblBody8 = .Controls.Add("Forms.Label.1", "DatepickerLblBody8")
      Set DatepickerLblBody9 = .Controls.Add("Forms.Label.1", "DatepickerLblBody9")
      Set DatepickerLblBody10 = .Controls.Add("Forms.Label.1", "DatepickerLblBody10")
      Set DatepickerLblBody11 = .Controls.Add("Forms.Label.1", "DatepickerLblBody11")
      Set DatepickerLblBody12 = .Controls.Add("Forms.Label.1", "DatepickerLblBody12")
      Set DatepickerLblBody13 = .Controls.Add("Forms.Label.1", "DatepickerLblBody13")
      Set DatepickerLblBody14 = .Controls.Add("Forms.Label.1", "DatepickerLblBody14")
      
      Set DatepickerLblBody15 = .Controls.Add("Forms.Label.1", "DatepickerLblBody15")
      Set DatepickerLblBody16 = .Controls.Add("Forms.Label.1", "DatepickerLblBody16")
      Set DatepickerLblBody17 = .Controls.Add("Forms.Label.1", "DatepickerLblBody17")
      Set DatepickerLblBody18 = .Controls.Add("Forms.Label.1", "DatepickerLblBody18")
      Set DatepickerLblBody19 = .Controls.Add("Forms.Label.1", "DatepickerLblBody19")
      Set DatepickerLblBody20 = .Controls.Add("Forms.Label.1", "DatepickerLblBody20")
      Set DatepickerLblBody21 = .Controls.Add("Forms.Label.1", "DatepickerLblBody21")
      
      Set DatepickerLblBody22 = .Controls.Add("Forms.Label.1", "DatepickerLblBody22")
      Set DatepickerLblBody23 = .Controls.Add("Forms.Label.1", "DatepickerLblBody23")
      Set DatepickerLblBody24 = .Controls.Add("Forms.Label.1", "DatepickerLblBody24")
      Set DatepickerLblBody25 = .Controls.Add("Forms.Label.1", "DatepickerLblBody25")
      Set DatepickerLblBody26 = .Controls.Add("Forms.Label.1", "DatepickerLblBody26")
      Set DatepickerLblBody27 = .Controls.Add("Forms.Label.1", "DatepickerLblBody27")
      Set DatepickerLblBody28 = .Controls.Add("Forms.Label.1", "DatepickerLblBody28")
      
      Set DatepickerLblBody29 = .Controls.Add("Forms.Label.1", "DatepickerLblBody29")
      Set DatepickerLblBody30 = .Controls.Add("Forms.Label.1", "DatepickerLblBody30")
      Set DatepickerLblBody31 = .Controls.Add("Forms.Label.1", "DatepickerLblBody31")
      Set DatepickerLblBody32 = .Controls.Add("Forms.Label.1", "DatepickerLblBody32")
      Set DatepickerLblBody33 = .Controls.Add("Forms.Label.1", "DatepickerLblBody33")
      Set DatepickerLblBody34 = .Controls.Add("Forms.Label.1", "DatepickerLblBody34")
      Set DatepickerLblBody35 = .Controls.Add("Forms.Label.1", "DatepickerLblBody35")
      
      Set DatepickerLblBody36 = .Controls.Add("Forms.Label.1", "DatepickerLblBody36")
      Set DatepickerLblBody37 = .Controls.Add("Forms.Label.1", "DatepickerLblBody37")
      Set DatepickerLblBody38 = .Controls.Add("Forms.Label.1", "DatepickerLblBody38")
      Set DatepickerLblBody39 = .Controls.Add("Forms.Label.1", "DatepickerLblBody39")
      Set DatepickerLblBody40 = .Controls.Add("Forms.Label.1", "DatepickerLblBody40")
      Set DatepickerLblBody41 = .Controls.Add("Forms.Label.1", "DatepickerLblBody41")
      Set DatepickerLblBody42 = .Controls.Add("Forms.Label.1", "DatepickerLblBody42")
      
      .Top = 35
      .Left = -1
      col = 9 '10
      row = 5 '40
      j = 1
      For i = 1 To 42
        With .Controls("DatepickerLblBody" & i)
          .Caption = i
          .Width = 12
          .Height = 12
          .TextAlign = fmTextAlignCenter
          .Top = row '10
          .Left = col  '(i * 12) + 5 ' (.Width + 4)
          .BackStyle = fmBackStyleTransparent
          .ForeColor = cDayForeColor ' VBA.RGB(100, 100, 100)
          .Tag = i
          
          If (i Mod 7) = 0 Then
            col = 9
            row = row + .Height * 1.15
            j = 1
            .ForeColor = VBA.RGB(255, 100, 100) ' VBA.vbRed
          Else
            col = 9 + (j * .Width * 1.15) 'col + 1
            j = j + 1
          End If
        End With
      Next i
    End With
  End With
    
  Call DatePicker_InicializaFormularioCalendario(Frm)
  cInstance = True

DP_SC:
  On Error GoTo 0
End Sub

Private Sub DatePicker_close(ByVal dFrm As Object)
  If Not DatepickerFrame2 Is Nothing Then
    Let dFrm.Controls("DatepickerFrame1").Visible = False
    Call dFrm.Controls.Remove("DatepickerFrame1")
  End If
  Call DatePicker_reset
End Sub

Private Sub DatePicker_reset()
  Set DatepickerFrame2 = Nothing
  
  Set DatepickerLblHead2 = Nothing
  
  Set DatepickerSpinHead1 = Nothing
  Set DatepickerSpinHead2 = Nothing
  
  Set DatepickerLblBody1 = Nothing
  Set DatepickerLblBody2 = Nothing
  Set DatepickerLblBody3 = Nothing
  Set DatepickerLblBody4 = Nothing
  Set DatepickerLblBody5 = Nothing
  Set DatepickerLblBody6 = Nothing
  Set DatepickerLblBody7 = Nothing
  
  Set DatepickerLblBody8 = Nothing
  Set DatepickerLblBody9 = Nothing
  Set DatepickerLblBody10 = Nothing
  Set DatepickerLblBody11 = Nothing
  Set DatepickerLblBody12 = Nothing
  Set DatepickerLblBody13 = Nothing
  Set DatepickerLblBody14 = Nothing
  
  Set DatepickerLblBody15 = Nothing
  Set DatepickerLblBody16 = Nothing
  Set DatepickerLblBody17 = Nothing
  Set DatepickerLblBody18 = Nothing
  Set DatepickerLblBody19 = Nothing
  Set DatepickerLblBody20 = Nothing
  Set DatepickerLblBody21 = Nothing
  
  Set DatepickerLblBody22 = Nothing
  Set DatepickerLblBody23 = Nothing
  Set DatepickerLblBody24 = Nothing
  Set DatepickerLblBody25 = Nothing
  Set DatepickerLblBody26 = Nothing
  Set DatepickerLblBody27 = Nothing
  Set DatepickerLblBody28 = Nothing
  
  Set DatepickerLblBody29 = Nothing
  Set DatepickerLblBody30 = Nothing
  Set DatepickerLblBody31 = Nothing
  Set DatepickerLblBody32 = Nothing
  Set DatepickerLblBody33 = Nothing
  Set DatepickerLblBody34 = Nothing
  Set DatepickerLblBody35 = Nothing
  
  Set DatepickerLblBody36 = Nothing
  Set DatepickerLblBody37 = Nothing
  Set DatepickerLblBody38 = Nothing
  Set DatepickerLblBody39 = Nothing
  Set DatepickerLblBody40 = Nothing
  Set DatepickerLblBody41 = Nothing
  Set DatepickerLblBody42 = Nothing
  
  cInstance = False
End Sub

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

Private Sub DatePicker_InicializaFormularioCalendario(ByVal dFrm As MSForms.UserForm)
  With dFrm
    .Controls("DatepickerFrame1").Controls("DatepickerLblHead1").Caption = "     " & VBA.StrConv(VBA.MonthName(cMonths), vbProperCase) & ", " & cYears
  End With
  With DatepickerFrame2.Controls("DatepickerLblBody" & cDays)
    DayBorderStyle = .BorderStyle
    DayBorderColor = .BorderColor
    DayBackColor = .BackColor
    DayForeColor = .ForeColor
  End With
  
  DatepickerSpinHead1.Value = VBA.Month(VBA.Date())
  DatepickerSpinHead2.Value = VBA.Year(VBA.Date())
  Call DatePicker_CargarLosDias(cYears, cMonths)
End Sub

Private Sub DatePicker_CargarLosDias(ByVal Ano As Long, ByVal Mes As Long)
  Dim FechaDelPrimerDia As Date
  Dim FechaDelUltimoDia As Date
  Dim DiaSemanaPrimerDia As Long
  Dim VariableControl As Control
  Dim Contador As Long

  FechaDelPrimerDia = VBA.DateSerial(Ano, Mes, 1)
  FechaDelUltimoDia = Application.WorksheetFunction.EoMonth(VBA.DateSerial(Ano, Mes, 1), 0)
  DiaSemanaPrimerDia = Application.WorksheetFunction.Weekday(FechaDelPrimerDia, 2)
  Contador = 1

  For Each VariableControl In DatepickerFrame2.Controls
    VariableControl.Visible = (VariableControl.Tag >= DiaSemanaPrimerDia And Contador <= VBA.Day(FechaDelUltimoDia))
    If (Contador = VBA.Day(VBA.Date())) And Not cMarkedDay Then
      cUnDays = Contador + DiaSemanaPrimerDia - 1
      cMarkedDay = True
    End If
    If VariableControl.Tag >= DiaSemanaPrimerDia And Contador <= VBA.Day(FechaDelUltimoDia) Then
      VariableControl.Caption = Contador
      Contador = Contador + 1
    End If
  Next VariableControl
End Sub
''********************************** NO MODIFICAR SI NO SABE **********************************
''*************************************|||||||||||||||||||*************************************
''***************************************|||||||||||||||***************************************
''*****************************************|||||||||||*****************************************
''*******************************************|||||||*******************************************
' ANDRÃ‰S ROJAS MONCADA'S CALENDAR CORE ALGORITHM                                ----- o ------


Private Sub DatePicker_ChangeMonthYear(ByVal cmValue As Long, Optional cmMonthYear As Boolean = False)
  With DatepickerSpinHead1
    If Not (VBA.IsNull(.Value)) Then
      If Not cmMonthYear Then 'Monts
        With DatepickerSpinHead1
          If cmValue > 12 Then
            .Value = 1: cYears = cYears + 1
            If cYears >= 9999 Then cYears = 9999
          End If
          If cmValue < 1 Then
            .Value = 12: cYears = cYears - 1
            If cYears <= 1900 Then cYears = 1900
          End If
          cMonths = .Value
          .Parent.Controls("DatepickerLblHead1").Caption = "     " & VBA.StrConv(VBA.MonthName(cMonths), vbProperCase) & ", " & cYears
        End With
      Else 'Years
        With DatepickerSpinHead2
          If cmValue >= 9999 Then
            cYears = 9999
          ElseIf cmValue <= 1900 Then
            cYears = 1900
          Else
            cYears = .Value
          End If
          .Parent.Controls("DatepickerLblHead1").Caption = "     " & VBA.StrConv(VBA.MonthName(cMonths), vbProperCase) & ", " & cYears
        End With
      End If

      Call DatePicker_CargarLosDias(cYears, cMonths)
      Call DatePicker_RemarkTheDay
    End If
  End With
End Sub

Private Sub DatePicker_RemarkTheDay()
  With DatepickerFrame2.Controls("DatepickerLblBody" & cUnDays)
    If DatepickerSpinHead1.Value = VBA.Month(VBA.Date()) And cYears = VBA.Year(VBA.Date()) Then
      .BorderStyle = fmBorderStyleSingle
      .BorderColor = DayBorderColor
      .BackColor = cDayBackColor
      .ForeColor = cDayForeColor
      .BackStyle = fmBackStyleOpaque
    Else
      .BorderStyle = DayBorderStyle
      .BorderColor = DayBorderColor
      .BackColor = DayBackColor
      .ForeColor = DayForeColor
      .BackStyle = fmBackStyleTransparent
    End If
  End With
End Sub

Private Sub DatePicker_RemarkSelectedDay(ByVal cLblDay As MSForms.Label)
  Dim ctrl As MSForms.Label

  For Each ctrl In DatepickerFrame2.Controls
    With ctrl
      .ForeColor = DayForeColor
      .BackStyle = fmBackStyleTransparent
    End With
  Next ctrl
  With DatepickerFrame2.Controls("DatepickerLblBody" & cUnDays)
    .ForeColor = DayForeColor
    .BackStyle = fmBackStyleTransparent
  End With
  With cLblDay
    .BackColor = cDayBackColor
    .ForeColor = cDayForeColor ' VBA.vbWhite
    .BackStyle = fmBackStyleOpaque
  End With
End Sub

Private Sub DatePicker_EventSelectDate(ByVal cLbl As MSForms.Label, Optional cToday As Boolean = False)
  cLbl.ForeColor = DayForeColor
  DatepickerFrame2.Parent.Visible = False 'FrameCal1.Visible = False
  CDateSelected = IIf(cToday, VBA.Date(), VBA.Format(cLbl.Caption, "00") & "/" & VBA.Format(cMonths, "00") & "/" & cYears)
  If cMsgBox Then Call VBA.MsgBox("Fecha: " & IIf(cToday, VBA.Date(), VBA.Format(cLbl.Caption, "00") & "/" & VBA.Format(cMonths, "00") & "/" & cYears))
  If Not cToday Then Call DatePicker_RemarkSelectedDay(cLbl)
  If Not cReceptorControl Is Nothing Then
    Select Case VBA.TypeName(cReceptorControl)
      Case "TextBox"
        cReceptorControl.Text = CDateSelected
      Case "ComboBox", "Range"
        cReceptorControl.Value = CDateSelected
      Case "Label"
        cReceptorControl.Caption = CDateSelected
    End Select
  End If
  If cJustForm Then
    Unload DatepickerFrame2.Parent.Parent
  Else
    Call DatePicker_close(DatepickerFrame2.Parent.Parent)
  End If
End Sub

Private Sub DatePicker_EventMouseMove(ByVal cLbl As MSForms.Label)
  Dim ctrl As MSForms.Label

  For Each ctrl In DatepickerFrame2.Controls
    ctrl.ForeColor = IIf((VBA.CLng(ctrl.Tag) Mod 7) = 0, DaySunForecolor, cDayForeColor)
  Next ctrl
  cLbl.ForeColor = IIf((VBA.CLng(cLbl.Tag) Mod 7) = 0, VBA.vbRed, VBA.vbBlack)
End Sub

Private Sub DatePicker_ChangeColors(ByVal FrameCal1 As MSForms.Frame)
  With FrameCal1
    .Controls("DatepickerLblHead1").ForeColor = cDayForeColor
    .Controls("DatepickerLblHead1").BackColor = cDayBackColor
  End With
End Sub

Private Sub DatePicker_Add(ByVal dFrm As MSForms.UserForm, Optional dFirstDay As Long = VBA.vbMonday)
    
  On Error Resume Next
  If dFrm Is Nothing Then GoTo DPA
  Call DatePicker_SetControls(dFrm, dFirstDay)

DPA:
  On Error GoTo 0
End Sub

Private Function DatePicker_Use( _
  ByVal TheForm As Object, _
  Optional clJustForm As Boolean = False, _
  Optional clControl As Object, _
  Optional clAlign As String = "Left", _
  Optional clBaseColor As Long = 4616993, _
  Optional clMsgBox As Boolean = False) As Date

  'Optional clControl As MSForms.Control,
  Dim fHeight As Single, fWidth As Single
  'Dim cc As Range
  'Dim tf As UserForm1
  
  DayBorderColor = clBaseColor
  cDayBackColor = clBaseColor
  cDayForeColor = clBaseColor + cDaySideColor 'Triade, soft color

  cJustForm = clJustForm
  cMsgBox = clMsgBox
  
  On Error Resume Next
  With TheForm
    If clJustForm Then
      .Controls("DatepickerFrame1").Left = 0
      .Controls("DatepickerFrame1").Top = 0
      .Controls("DatepickerFrame1").Visible = True ' Not .Controls("DatepickerFrame1").Visible
      fWidth = .Controls("DatepickerFrame1").Width
      fHeight = .Controls("DatepickerFrame1").Height
      .Width = fWidth * 1.09
      .Height = fHeight * 1.2
      If VBA.TypeName(clControl) = "Range" Then
        .StartUpPosition = 0
        .Left = clControl.Left
        .Top = .Height + clControl.Top + clControl.Height
      End If
      Call DatePicker_ChangeColors(.Controls("DatepickerFrame1"))
      .Caption = "DatePickerClass"
      Load TheForm
      .Show
    Else
      If Not clControl Is Nothing Then
        Select Case VBA.LCase(clAlign)
          Case "right"
            .Controls("DatepickerFrame1").Left = (clControl.Left + clControl.Width) - .Controls("DatepickerFrame1").Width
          Case "center"
            .Controls("DatepickerFrame1").Left = (clControl.Left + (clControl.Width / 2)) - (.Controls("DatepickerFrame1").Width / 2)
          Case Else
            .Controls("DatepickerFrame1").Left = clControl.Left
        End Select
        Set cReceptorControl = clControl
        .Controls("DatepickerFrame1").Top = clControl.Top + clControl.Height ' + 2
        .Controls("DatepickerFrame1").Visible = True 'Not FrameCal1.Visible
        Call DatePicker_ChangeColors(.Controls("DatepickerFrame1"))
      End If
    End If
    If Not DatepickerFrame2 Is Nothing Then .Controls("DatepickerFrame1").Visible = True
  End With
  DatePicker_Use = CDateSelected
End Function





'METHODS PUBLIC
'Interface DatePickerClass
Public Function DatePicker( _
  ByVal Frm As Object, _
  Optional dJustForm As Boolean = False, _
  Optional dFirstDay As Long = VBA.vbMonday, _
  Optional dControl As Object, _
  Optional dAlign As String = "Left", _
  Optional dBaseColor As Long = 4616993, _
  Optional dMsgBox As Boolean = False) As Date
  
  DatePicker = VBA.Format(VBA.Date(), "dd/mm/yy")
  If Not Frm Is Nothing Then
    If cInstance Then Call DatePicker_close(Frm): GoTo dp
    Call DatePicker_Add(Frm, dFirstDay)
    DatePicker = DatePicker_Use(Frm, dJustForm, dControl, dAlign, dBaseColor, dMsgBox)
  End If

dp:
  On Error GoTo 0
End Function

