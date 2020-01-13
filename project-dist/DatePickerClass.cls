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
' Version: 1.0.0
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
' VersiÃ³n 1.0 - 20 de julio de 2015
' URL del canal: www.youtube.com/jarmoncada01
' Para ver algunos ejemplos sobre el uso de este calendario, observa este video.
' Enlace: |||||||||||||||||||| https://www.youtube.com/watch?v=FkjsuN2zqSU ||||||||||||||||||||
' ! Muchas gracias y espero lo disfruten ! ---------------------------------------------------
' ------------------------------------------------------------------------------------------------------------

Private TheForm As MSForms.UserForm

Private WithEvents FrameCal1 As MSForms.Frame
Attribute FrameCal1.VB_VarHelpID = -1
Private FrameCal2 As MSForms.Frame
Attribute FrameCal2.VB_VarHelpID = -1

Private WithEvents BtnCal1 As MSForms.CommandButton
Attribute BtnCal1.VB_VarHelpID = -1

Private LblCalHead1 As MSForms.Label
Attribute LblCalHead1.VB_VarHelpID = -1
Private WithEvents LblCalHead2 As MSForms.Label
Attribute LblCalHead2.VB_VarHelpID = -1
Private LblCalHead30 As MSForms.Label
Attribute LblCalHead30.VB_VarHelpID = -1
Private LblCalHead31 As MSForms.Label
Private LblCalHead32 As MSForms.Label
Private LblCalHead33 As MSForms.Label
Private LblCalHead34 As MSForms.Label
Private LblCalHead35 As MSForms.Label
Private LblCalHead36 As MSForms.Label

Private WithEvents SpinCalHead1 As MSForms.SpinButton
Attribute SpinCalHead1.VB_VarHelpID = -1

Private WithEvents LblCal1 As MSForms.Label
Attribute LblCal1.VB_VarHelpID = -1
Private WithEvents LblCal2 As MSForms.Label
Attribute LblCal2.VB_VarHelpID = -1
Private WithEvents LblCal3 As MSForms.Label
Attribute LblCal3.VB_VarHelpID = -1
Private WithEvents LblCal4 As MSForms.Label
Attribute LblCal4.VB_VarHelpID = -1
Private WithEvents LblCal5 As MSForms.Label
Attribute LblCal5.VB_VarHelpID = -1
Private WithEvents LblCal6 As MSForms.Label
Attribute LblCal6.VB_VarHelpID = -1
Private WithEvents LblCal7 As MSForms.Label
Attribute LblCal7.VB_VarHelpID = -1

Private WithEvents LblCal8 As MSForms.Label
Attribute LblCal8.VB_VarHelpID = -1
Private WithEvents LblCal9 As MSForms.Label
Attribute LblCal9.VB_VarHelpID = -1
Private WithEvents LblCal10 As MSForms.Label
Attribute LblCal10.VB_VarHelpID = -1
Private WithEvents LblCal11 As MSForms.Label
Attribute LblCal11.VB_VarHelpID = -1
Private WithEvents LblCal12 As MSForms.Label
Attribute LblCal12.VB_VarHelpID = -1
Private WithEvents LblCal13 As MSForms.Label
Attribute LblCal13.VB_VarHelpID = -1
Private WithEvents LblCal14 As MSForms.Label
Attribute LblCal14.VB_VarHelpID = -1

Private WithEvents LblCal15 As MSForms.Label
Attribute LblCal15.VB_VarHelpID = -1
Private WithEvents LblCal16 As MSForms.Label
Attribute LblCal16.VB_VarHelpID = -1
Private WithEvents LblCal17 As MSForms.Label
Attribute LblCal17.VB_VarHelpID = -1
Private WithEvents LblCal18 As MSForms.Label
Attribute LblCal18.VB_VarHelpID = -1
Private WithEvents LblCal19 As MSForms.Label
Attribute LblCal19.VB_VarHelpID = -1
Private WithEvents LblCal20 As MSForms.Label
Attribute LblCal20.VB_VarHelpID = -1
Private WithEvents LblCal21 As MSForms.Label
Attribute LblCal21.VB_VarHelpID = -1

Private WithEvents LblCal22 As MSForms.Label
Attribute LblCal22.VB_VarHelpID = -1
Private WithEvents LblCal23 As MSForms.Label
Attribute LblCal23.VB_VarHelpID = -1
Private WithEvents LblCal24 As MSForms.Label
Attribute LblCal24.VB_VarHelpID = -1
Private WithEvents LblCal25 As MSForms.Label
Attribute LblCal25.VB_VarHelpID = -1
Private WithEvents LblCal26 As MSForms.Label
Attribute LblCal26.VB_VarHelpID = -1
Private WithEvents LblCal27 As MSForms.Label
Attribute LblCal27.VB_VarHelpID = -1
Private WithEvents LblCal28 As MSForms.Label
Attribute LblCal28.VB_VarHelpID = -1

Private WithEvents LblCal29 As MSForms.Label
Attribute LblCal29.VB_VarHelpID = -1
Private WithEvents LblCal30 As MSForms.Label
Attribute LblCal30.VB_VarHelpID = -1
Private WithEvents LblCal31 As MSForms.Label
Attribute LblCal31.VB_VarHelpID = -1
Private WithEvents LblCal32 As MSForms.Label
Attribute LblCal32.VB_VarHelpID = -1
Private WithEvents LblCal33 As MSForms.Label
Attribute LblCal33.VB_VarHelpID = -1
Private WithEvents LblCal34 As MSForms.Label
Attribute LblCal34.VB_VarHelpID = -1
Private WithEvents LblCal35 As MSForms.Label
Attribute LblCal35.VB_VarHelpID = -1

Private WithEvents LblCal36 As MSForms.Label
Attribute LblCal36.VB_VarHelpID = -1
Private WithEvents LblCal37 As MSForms.Label
Attribute LblCal37.VB_VarHelpID = -1
Private WithEvents LblCal38 As MSForms.Label
Attribute LblCal38.VB_VarHelpID = -1
Private WithEvents LblCal39 As MSForms.Label
Attribute LblCal39.VB_VarHelpID = -1
Private WithEvents LblCal40 As MSForms.Label
Attribute LblCal40.VB_VarHelpID = -1
Private WithEvents LblCal41 As MSForms.Label
Attribute LblCal41.VB_VarHelpID = -1
Private WithEvents LblCal42 As MSForms.Label
Attribute LblCal42.VB_VarHelpID = -1


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
  Set FrameCal1 = Nothing
  Set FrameCal2 = Nothing
    
  Set LblCalHead1 = Nothing
  Set LblCalHead2 = Nothing
  Set LblCalHead30 = Nothing
  Set LblCalHead31 = Nothing
  Set LblCalHead32 = Nothing
  Set LblCalHead33 = Nothing
  Set LblCalHead34 = Nothing
  Set LblCalHead35 = Nothing
  Set LblCalHead36 = Nothing
  
  Set SpinCalHead1 = Nothing
  
  Set LblCal1 = Nothing
  Set LblCal2 = Nothing
  Set LblCal3 = Nothing
  Set LblCal4 = Nothing
  Set LblCal5 = Nothing
  Set LblCal6 = Nothing
  Set LblCal7 = Nothing
  
  Set LblCal8 = Nothing
  Set LblCal9 = Nothing
  Set LblCal10 = Nothing
  Set LblCal11 = Nothing
  Set LblCal12 = Nothing
  Set LblCal13 = Nothing
  Set LblCal14 = Nothing
  
  Set LblCal15 = Nothing
  Set LblCal16 = Nothing
  Set LblCal17 = Nothing
  Set LblCal18 = Nothing
  Set LblCal19 = Nothing
  Set LblCal20 = Nothing
  Set LblCal21 = Nothing
  
  Set LblCal22 = Nothing
  Set LblCal23 = Nothing
  Set LblCal24 = Nothing
  Set LblCal25 = Nothing
  Set LblCal26 = Nothing
  Set LblCal27 = Nothing
  Set LblCal28 = Nothing
  
  Set LblCal29 = Nothing
  Set LblCal30 = Nothing
  Set LblCal31 = Nothing
  Set LblCal32 = Nothing
  Set LblCal33 = Nothing
  Set LblCal34 = Nothing
  Set LblCal35 = Nothing
  
  Set LblCal36 = Nothing
  Set LblCal37 = Nothing
  Set LblCal38 = Nothing
  Set LblCal39 = Nothing
  Set LblCal40 = Nothing
  Set LblCal41 = Nothing
  Set LblCal42 = Nothing
  
  Set cReceptorControl = Nothing
  Set TheForm = Nothing
End Sub



'METHODS PRIVATE
'Methods Controls Events
'Methods calendar labels picker Events
Private Sub SpinCalHead1_Change()
  Call DatePicker_ChangeMonthYear(SpinCalHead1.Value)
End Sub

Private Sub LblCalHead2_Click()
  Call DatePicker_EventSelectDate(LblCalHead2, True)
  LblCalHead2.ForeColor = DayForeColor
End Sub

Private Sub LblCalHead2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  LblCalHead2.ForeColor = cDayBackColor
End Sub


Private Sub LblCal1_Click()
  Call DatePicker_EventSelectDate(LblCal1)
End Sub

Private Sub LblCal1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal1)
End Sub

Private Sub LblCal2_Click()
  Call DatePicker_EventSelectDate(LblCal2)
End Sub

Private Sub LblCal2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal2)
End Sub

Private Sub LblCal3_Click()
  Call DatePicker_EventSelectDate(LblCal3)
End Sub

Private Sub LblCal3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal3)
End Sub

Private Sub LblCal4_Click()
  Call DatePicker_EventSelectDate(LblCal4)
End Sub

Private Sub LblCal4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal4)
End Sub

Private Sub LblCal5_Click()
  Call DatePicker_EventSelectDate(LblCal5)
End Sub

Private Sub LblCal5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal5)
End Sub

Private Sub LblCal6_Click()
  Call DatePicker_EventSelectDate(LblCal6)
End Sub

Private Sub LblCal6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal6)
End Sub

Private Sub LblCal7_Click()
  Call DatePicker_EventSelectDate(LblCal7)
End Sub

Private Sub LblCal7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal7)
End Sub

Private Sub LblCal8_Click()
  Call DatePicker_EventSelectDate(LblCal8)
End Sub

Private Sub LblCal8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal8)
End Sub

Private Sub LblCal9_Click()
  Call DatePicker_EventSelectDate(LblCal9)
End Sub

Private Sub LblCal9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal9)
End Sub

Private Sub LblCal10_Click()
  Call DatePicker_EventSelectDate(LblCal10)
End Sub

Private Sub LblCal10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal10)
End Sub

Private Sub LblCal11_Click()
  Call DatePicker_EventSelectDate(LblCal11)
End Sub

Private Sub LblCal11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal11)
End Sub

Private Sub LblCal12_Click()
  Call DatePicker_EventSelectDate(LblCal12)
End Sub

Private Sub LblCal12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal12)
End Sub

Private Sub LblCal13_Click()
  Call DatePicker_EventSelectDate(LblCal13)
End Sub

Private Sub LblCal13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal13)
End Sub

Private Sub LblCal14_Click()
  Call DatePicker_EventSelectDate(LblCal14)
End Sub

Private Sub LblCal14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal14)
End Sub

Private Sub LblCal15_Click()
  Call DatePicker_EventSelectDate(LblCal15)
End Sub

Private Sub LblCal15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal15)
End Sub

Private Sub LblCal16_Click()
  Call DatePicker_EventSelectDate(LblCal16)
End Sub

Private Sub LblCal16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal16)
End Sub

Private Sub LblCal17_Click()
  Call DatePicker_EventSelectDate(LblCal17)
End Sub

Private Sub LblCal17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal17)
End Sub

Private Sub LblCal18_Click()
  Call DatePicker_EventSelectDate(LblCal18)
End Sub

Private Sub LblCal18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal18)
End Sub

Private Sub LblCal19_Click()
  Call DatePicker_EventSelectDate(LblCal19)
End Sub

Private Sub LblCal19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal19)
End Sub

Private Sub LblCal20_Click()
  Call DatePicker_EventSelectDate(LblCal20)
End Sub

Private Sub LblCal20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal20)
End Sub

Private Sub LblCal21_Click()
  Call DatePicker_EventSelectDate(LblCal21)
End Sub

Private Sub LblCal21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal21)
End Sub

Private Sub LblCal22_Click()
  Call DatePicker_EventSelectDate(LblCal22)
End Sub

Private Sub LblCal22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal22)
End Sub

Private Sub LblCal23_Click()
  Call DatePicker_EventSelectDate(LblCal23)
End Sub

Private Sub LblCal23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal23)
End Sub

Private Sub LblCal24_Click()
  Call DatePicker_EventSelectDate(LblCal24)
End Sub

Private Sub LblCal24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal24)
End Sub

Private Sub LblCal25_Click()
  Call DatePicker_EventSelectDate(LblCal25)
End Sub

Private Sub LblCal25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal25)
End Sub

Private Sub LblCal26_Click()
  Call DatePicker_EventSelectDate(LblCal26)
End Sub

Private Sub LblCal26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal26)
End Sub

Private Sub LblCal27_Click()
  Call DatePicker_EventSelectDate(LblCal27)
End Sub

Private Sub LblCal27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal27)
End Sub

Private Sub LblCal28_Click()
  Call DatePicker_EventSelectDate(LblCal28)
End Sub

Private Sub LblCal28_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal28)
End Sub

Private Sub LblCal29_Click()
  Call DatePicker_EventSelectDate(LblCal29)
End Sub

Private Sub LblCal29_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal29)
End Sub

Private Sub LblCal30_Click()
  Call DatePicker_EventSelectDate(LblCal30)
End Sub

Private Sub LblCal30_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal30)
End Sub

Private Sub LblCal31_Click()
  Call DatePicker_EventSelectDate(LblCal31)
End Sub

Private Sub LblCal31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal31)
End Sub

Private Sub LblCal32_Click()
  Call DatePicker_EventSelectDate(LblCal32)
End Sub

Private Sub LblCal32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal32)
End Sub

Private Sub LblCal33_Click()
  Call DatePicker_EventSelectDate(LblCal33)
End Sub

Private Sub LblCal33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal33)
End Sub

Private Sub LblCal34_Click()
  Call DatePicker_EventSelectDate(LblCal34)
End Sub

Private Sub LblCal34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal34)
End Sub

Private Sub LblCal35_Click()
  Call DatePicker_EventSelectDate(LblCal35)
End Sub

Private Sub LblCal35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal35)
End Sub

Private Sub LblCal36_Click()
  Call DatePicker_EventSelectDate(LblCal36)
End Sub

Private Sub LblCal36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal36)
End Sub

Private Sub LblCal37_Click()
  Call DatePicker_EventSelectDate(LblCal37)
End Sub

Private Sub LblCal37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal37)
End Sub

Private Sub LblCal38_Click()
  Call DatePicker_EventSelectDate(LblCal38)
End Sub

Private Sub LblCal38_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal38)
End Sub

Private Sub LblCal39_Click()
  Call DatePicker_EventSelectDate(LblCal39)
End Sub

Private Sub LblCal39_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal39)
End Sub

Private Sub LblCal40_Click()
  Call DatePicker_EventSelectDate(LblCal40)
End Sub

Private Sub LblCal40_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal40)
End Sub

Private Sub LblCal41_Click()
  Call DatePicker_EventSelectDate(LblCal41)
End Sub

Private Sub LblCal41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblCal41)
End Sub

Private Sub LblCal42_Click()
  Call DatePicker_EventSelectDate(LblCal42)
End Sub

Private Sub LblCal42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call DatePicker_EventMouseMove(LblC42l5)
End Sub


'Methods Core Class
Private Sub DatePicker_SetControls(ByVal Frm As MSForms.UserForm, Optional cFirstDay As Long = VBA.vbMonday)
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
  
  With Frm
    Set FrameCal1 = .Controls.Add("Forms.Frame.1", "FrameCal1", False)
    With FrameCal1
      .Width = 114
      .Height = 130
      .BorderStyle = fmBorderStyleSingle
    End With
    
    Set LblCalHead1 = FrameCal1.Controls.Add("Forms.Label.1", "LblCalHead1")
    Set LblCalHead2 = FrameCal1.Controls.Add("Forms.Label.1", "LblCalHead2")
    
    Set LblCalHead30 = FrameCal1.Controls.Add("Forms.Label.1", "LblCalHead30")
    Set LblCalHead31 = FrameCal1.Controls.Add("Forms.Label.1", "LblCalHead31")
    Set LblCalHead32 = FrameCal1.Controls.Add("Forms.Label.1", "LblCalHead32")
    Set LblCalHead33 = FrameCal1.Controls.Add("Forms.Label.1", "LblCalHead33")
    Set LblCalHead34 = FrameCal1.Controls.Add("Forms.Label.1", "LblCalHead34")
    Set LblCalHead35 = FrameCal1.Controls.Add("Forms.Label.1", "LblCalHead35")
    Set LblCalHead36 = FrameCal1.Controls.Add("Forms.Label.1", "LblCalHead36")
    
    Set SpinCalHead1 = FrameCal1.Controls.Add("Forms.SpinButton.1", "SpinCalHead1")
    
    Set FrameCal2 = FrameCal1.Controls.Add("Forms.Frame.1", "FrameCal2")
    
    With SpinCalHead1
      .Height = 12
      .Width = FrameCal1.Width * 0.2
      .Left = FrameCal1.Width - (SpinCalHead1.Width + 10)
      .Delay = 50
      .Min = 0
      .Max = 13
    End With
    
    LblCalHead1.Height = 12
    LblCalHead1.Width = FrameCal1.Width
    LblCalHead1.Caption = "     Enero, 2020"
    LblCalHead1.Font.Bold = True
    
    LblCalHead2.Caption = "     Hoy: " & VBA.Date
    LblCalHead2.Top = 12
    LblCalHead2.BackStyle = fmBackStyleTransparent
    LblCalHead2.ForeColor = DaysColor ' VBA.RGB(100, 100, 100)
    
    For i = 0 To 6
      With FrameCal1.Controls("LblCalHead3" & i)
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
    
    Set LblCal1 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal1")
    Set LblCal2 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal2")
    Set LblCal3 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal3")
    Set LblCal4 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal4")
    Set LblCal5 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal5")
    Set LblCal6 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal6")
    Set LblCal7 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal7")
    
    Set LblCal8 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal8")
    Set LblCal9 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal9")
    Set LblCal10 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal10")
    Set LblCal11 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal11")
    Set LblCal12 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal12")
    Set LblCal13 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal13")
    Set LblCal14 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal14")
    
    Set LblCal15 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal15")
    Set LblCal16 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal16")
    Set LblCal17 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal17")
    Set LblCal18 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal18")
    Set LblCal19 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal19")
    Set LblCal20 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal20")
    Set LblCal21 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal21")
    
    Set LblCal22 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal22")
    Set LblCal23 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal23")
    Set LblCal24 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal24")
    Set LblCal25 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal25")
    Set LblCal26 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal26")
    Set LblCal27 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal27")
    Set LblCal28 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal28")
    
    Set LblCal29 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal29")
    Set LblCal30 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal30")
    Set LblCal31 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal31")
    Set LblCal32 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal32")
    Set LblCal33 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal33")
    Set LblCal34 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal34")
    Set LblCal35 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal35")
    
    Set LblCal36 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal36")
    Set LblCal37 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal37")
    Set LblCal38 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal38")
    Set LblCal39 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal39")
    Set LblCal40 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal40")
    Set LblCal41 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal41")
    Set LblCal42 = FrameCal2.Controls.Add("Forms.Label.1", "LblCal42")
    
    FrameCal2.Top = 35
    FrameCal2.Left = -1
    col = 9 '10
    row = 5 '40
    j = 1
    For i = 1 To 42
      With FrameCal2.Controls("LblCal" & i)
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
    
  Call DatePicker_InicializaFormularioCalendario

DP_SC:
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

Private Sub DatePicker_InicializaFormularioCalendario()
  LblCalHead1.Caption = "     " & VBA.StrConv(VBA.MonthName(cMonths), vbProperCase) & ", " & cYears
  
  With FrameCal2.Controls("LblCal" & cDays)
    DayBorderStyle = .BorderStyle
    DayBorderColor = .BorderColor
    DayBackColor = .BackColor
    DayForeColor = .ForeColor
  End With
  
  SpinCalHead1.Value = VBA.Month(VBA.Date())
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

  For Each VariableControl In FrameCal2.Controls
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


Private Sub DatePicker_ChangeMonthYear(ByVal cmMonth As Long)
    If Not (VBA.IsNull(SpinCalHead1.Value)) Then
      If cmMonth > 12 Then
        SpinCalHead1.Value = 1: cYears = cYears + 1
        If cYears >= 9999 Then cYears = 9999
      End If
      If cmMonth < 1 Then
        SpinCalHead1.Value = 12: cYears = cYears - 1
        If cYears <= 1900 Then cYears = 1900
      End If
      
      cMonths = SpinCalHead1.Value
      LblCalHead1.Caption = "     " & VBA.StrConv(VBA.MonthName(cMonths), vbProperCase) & ", " & cYears
      Call DatePicker_CargarLosDias(cYears, cMonths)
      Call DatePicker_RemarkTheDay
    End If
End Sub

Private Sub DatePicker_RemarkTheDay()
  With FrameCal2.Controls("LblCal" & cUnDays)
    If SpinCalHead1.Value = VBA.Month(VBA.Date()) And cYears = VBA.Year(VBA.Date()) Then
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

  For Each ctrl In FrameCal2.Controls
    With ctrl
      .ForeColor = DayForeColor
      .BackStyle = fmBackStyleTransparent
    End With
  Next ctrl
  With FrameCal2.Controls("LblCal" & cUnDays)
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
  FrameCal1.Visible = False
  CDateSelected = IIf(cToday, VBA.Date(), VBA.Format(cLbl.Caption, "00") & "/" & VBA.Format(cMonths, "00") & "/" & cYears)
  If cMsgBox Then Call VBA.MsgBox("Fecha: " & IIf(cToday, VBA.Date(), VBA.Format(cLbl.Caption, "00") & "/" & VBA.Format(cMonths, "00") & "/" & cYears))
  If Not cToday Then Call DatePicker_RemarkSelectedDay(cLbl)
  If Not cReceptorControl Is Nothing Then
    Debug.Print VBA.TypeName(cReceptorControl)
    Select Case VBA.TypeName(cReceptorControl)
      Case "TextBox"
        cReceptorControl.Text = CDateSelected
      Case "ComboBox"
        cReceptorControl.Value = CDateSelected
      Case "Label"
        cReceptorControl.Caption = CDateSelected
    End Select
  End If
  If cJustForm Then Unload TheForm
End Sub

Private Sub DatePicker_EventMouseMove(ByVal cLbl As MSForms.Label)
  Dim ctrl As MSForms.Label
  
  For Each ctrl In FrameCal2.Controls
    ctrl.ForeColor = IIf((VBA.CLng(ctrl.Tag) Mod 7) = 0, DaySunForecolor, cDayForeColor)
  Next ctrl
  cLbl.ForeColor = IIf((VBA.CLng(cLbl.Tag) Mod 7) = 0, VBA.vbRed, VBA.vbBlack)
End Sub

Private Sub DatePicker_ChangeColors()
  With FrameCal1
    .Controls("LblCalHead1").ForeColor = cDayForeColor
    .Controls("LblCalHead1").BackColor = cDayBackColor
  End With
End Sub




'METHODS PUBLIC
'Interface DatePickerClass
Public Sub DatePickerAdd(ByRef clFrm As MSForms.UserForm, Optional clFirstDay As Long = VBA.vbMonday)
    
  If clFrm Is Nothing Then GoTo DPA
  Set TheForm = clFrm
  On Error Resume Next
  If Not cInstance Then 'Just one instance
    With TheForm
      Call DatePicker_SetControls(clFrm, clFirstDay)
      cInstance = True
    End With
  End If

DPA:
End Sub

Public Function DatePickerUse( _
  Optional clJustForm As Boolean = False, _
  Optional clControl As MSForms.Control, _
  Optional clAlign As String = "L", _
  Optional clBaseColor As Long = 4616993, _
  Optional clMsgBox As Boolean = False) As Date
  
  DayBorderColor = clBaseColor
  cDayBackColor = clBaseColor
  cDayForeColor = clBaseColor + cDaySideColor 'Triade, soft color
  
  cJustForm = clJustForm
  cMsgBox = clMsgBox
  On Error Resume Next
  If clJustForm Then
    If Not TheForm Is Nothing Then
      With TheForm
        .Controls("FrameCal1").Left = 0
        .Controls("FrameCal1").Top = 0
        .Controls("FrameCal1").Visible = Not .Controls("FrameCal1").Visible
        'Debug.Print .Controls("FrameCal1").Width
        'Debug.Print .Controls("FrameCal1").Height
        .Width = .Controls("FrameCal1").Width
        .Height = .Controls("FrameCal1").Height
      End With
      Call DatePicker_ChangeColors
      'TheForm.Show
    End If
  Else
    If Not clControl Is Nothing Then
      Select Case clAlign
        Case "R"
          FrameCal1.Left = (clControl.Left + clControl.Width) - FrameCal1.Width
        Case "C"
          FrameCal1.Left = (clControl.Left + (clControl.Width / 2)) - (FrameCal1.Width / 2)
        Case Else
          FrameCal1.Left = clControl.Left
      End Select
      Set cReceptorControl = clControl
      FrameCal1.Top = clControl.Top + clControl.Height ' + 2
      FrameCal1.Visible = Not FrameCal1.Visible
      Call DatePicker_ChangeColors
    End If
  End If
  DatePickerUse = CDateSelected
End Function
