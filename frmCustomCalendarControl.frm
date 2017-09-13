VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCustomCalendarControl 
   Caption         =   "Select Date..."
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4020
   OleObjectBlob   =   "frmCustomCalendarControl.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmCustomCalendarControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Private mSelectedDay As Integer
Private mSelectedMonth As Integer
Private mSelectedYear As Integer
Private OriginalDay As Integer
Private OriginalMonth As Integer
Private OriginalYear As Integer
Private ComboBoxBeingUpdatedBySystem As Boolean
Private IgnoreEvents As Boolean
'===============================================================================================================================
'=
'= VBA CUSTOM CALENDAR CONTROL
'=
'===============================================================================================================================
'=
'= Developed By : Kevin Clark, Risk Tools Team, 29/04/2009
'=
'= This custom calendar control has been developed to be a direct replacement for the calendar
'= control which is built into versions of Excel 2003 and later. The look and feel of this custom
'= calendar control is intentionally identical to the look and feel of the calendar control provided
'= by Microsoft.
'=
'= This custom calendar control has been developed due to the lack of consistent availability of
'= a Microsoft provided calendar control across all versions of Excel, particularly Excel 2000
'= which dependant upon the type of Barclays build in use can have no calendar control at all
'= available. This makes it extremely difficult for a developer to use a calendar control in an Excel
'= based applciation and to guarantee the required level of stability of the system across all
'= platforms and builds.
'=
'= This custom calendar control is based entirely in bespoke VBA code. The custom calendar control
'= is fully contained within a single form and has no dependancies on any other components
'= whatsoever. To use this all you need to do is to import the form frmCustomCalendarControl into
'= your VBA project and then use it in much the same way as the frmTestHarness form in this example
'= already does.
'=
'===============================================================================================================================
'=
'= The custom calendar control has the following properties available :-
'=
'= Property Name         Access      Type     Summary
'= -------------------   ----------  -------  ----------------------------------------------------------------------------------
'= SelectedDayNumber     Read/Write  Integer  A number from 1 to 31 indicating the day selected in the calendar control
'= SelectedMonthNumber   Read/Write  Integer  A number from 1 to 12 indicating the month selected in the calendar control
'= SelectedYearNumber    Read/Write  Integer  A number from 1901 to 2199 indicating the year selected in the calendar control
'= SelectedDateDDMMYYYY  Read Only   String   The date in DDMMYYYY format that has been selected in the calendar control
'= SelectedDayString     Read Only   String   The day of the wek that has been selected in the calendar control, e.g. "Monday"
'= SelectedMonthString   Read Only   String   The month that has been selected in the calendar control, e.g. "January"
'= SelectedYearString    ReadOnly    String   Teh year that has been selected in the calendar control, returned as a string
'=
'===============================================================================================================================


Public Property Get SelectedDayNumber() As Integer
  SelectedDayNumber = mSelectedDay
End Property
Public Property Get SelectedMonthNumber() As Integer
  SelectedMonthNumber = mSelectedMonth
End Property
Public Property Get SelectedYearNumber() As Integer
  SelectedYearNumber = mSelectedYear
End Property
Public Property Let SelectedDayNumber(IncomingDayNumber As Integer)
  If IncomingDayNumber < 1 Or IncomingDayNumber > 31 Then
     ' an invalid day is being passed in as property so assume that it is the current day instead
     mSelectedDay = Day(Now())
  Else
     mSelectedDay = IncomingDayNumber
  End If
  OriginalDay = mSelectedDay
End Property
Public Property Let SelectedMonthNumber(IncomingMonthNumber As Integer)
  If IncomingMonthNumber < 1 Or IncomingMonthNumber > 12 Then
     ' an invalid month is being passed in as property so assume that it is the current month instead
     mSelectedMonth = Month(Now())
  Else
     mSelectedMonth = IncomingMonthNumber
  End If
  OriginalMonth = mSelectedMonth
End Property
Public Property Let SelectedYearNumber(IncomingYearNumber As Integer)
  If IncomingYearNumber < 1901 Or IncomingYearNumber > 2199 Then
     ' an invalid year is being passed in as property so assume that it is the current year instead
     mSelectedYear = Year(Now())
  Else
     mSelectedYear = IncomingYearNumber
  End If
  OriginalYear = mSelectedYear
End Property
Public Property Get SelectedDateDDMMYYYY() As String
  SelectedDateDDMMYYYY = Format(mSelectedDay, "00") & _
                         "/" & _
                         Format(mSelectedMonth, "00") & _
                         "/" & _
                         Format(mSelectedYear, "0000")
End Property
Public Property Get SelectedDayString() As String
  SelectedDayString = WeekdayName(Weekday(SelectedDateDDMMYYYY), False, vbSunday)
End Property
Public Property Get SelectedMonthString() As String
  SelectedMonthString = MonthName(Month(SelectedDateDDMMYYYY))
End Property
Public Property Get SelectedYearString() As String
  SelectedYearString = Trim(CStr(Year(SelectedDateDDMMYYYY)))
End Property
Private Sub UserForm_Activate()

  Dim idxMonth As Integer
  Dim idxYear As Integer

  ' initialise the flag used to ignore event handling
  IgnoreEvents = False

  ' populate all month names into the month combo box
  ComboBoxBeingUpdatedBySystem = True
  cmbMonth.Clear
  For idxMonth = 1 To 12
      cmbMonth.AddItem (MonthName(idxMonth))
  Next idxMonth
  ComboBoxBeingUpdatedBySystem = False

  ' populate all years into the year combo box
  ComboBoxBeingUpdatedBySystem = True
  cmbYear.Clear
  For idxYear = 1901 To 2199
      cmbYear.AddItem (Trim(CStr(idxYear)))
  Next idxYear
  ComboBoxBeingUpdatedBySystem = False
  
  ' if no date has been passed into this form then assume that the default selected date should
  ' be today's date
  If mSelectedDay = 0 Then mSelectedDay = Day(Now())
  If mSelectedMonth = 0 Then mSelectedMonth = Month(Now())
  If mSelectedYear = 0 Then mSelectedYear = Year(Now())
  
  ' select correct month in month combo box
  ComboBoxBeingUpdatedBySystem = True
  cmbMonth.ListIndex = (mSelectedMonth - 1) ' combo box list index is zero bound
  ComboBoxBeingUpdatedBySystem = False
  
  ' select correct year in year combo box
  ComboBoxBeingUpdatedBySystem = True
  cmbYear.Text = mSelectedYear
  ComboBoxBeingUpdatedBySystem = False
  
  ' redraw the day toggle buttons to reflect the currently selected month and year
  Call RedrawDayToggleButtons
  
  ' select default day in array of day toggle buttons
  Call SelectDefaultDay
  
  ' if the current user is a developer then show the label used to display the selected date
  If UserIsADeveloper() Then
     lblDebugDate.Visible = True
  Else
     lblDebugDate.Visible = False
  End If

  ' put focus on the Cancel button
  cmdCancel.SetFocus

End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

  ' the user has to click on the OK button if they want to close the form, not the X button
  Cancel = True
 
End Sub
Private Sub cmbMonth_Change()

  If ComboBoxBeingUpdatedBySystem = True Then
     ' do nothing
  Else
     mSelectedMonth = (cmbMonth.ListIndex + 1) ' month combo box is zero bound
     Call RedrawDayToggleButtons
  End If
  
End Sub
Private Sub cmbYear_Change()

  If ComboBoxBeingUpdatedBySystem = True Then
     ' do nothing
  Else
     mSelectedYear = CInt((cmbYear.Text))
     Call RedrawDayToggleButtons
  End If
  
End Sub
Private Sub cmdCancel_Click()

  ' recover the original date that was passed into this form
  mSelectedDay = OriginalDay
  mSelectedMonth = OriginalMonth
  mSelectedYear = OriginalYear
  
  ' hide the form
  frmCustomCalendarControl.Hide

End Sub
Private Sub cmdOK_Click()

  ' hide the form
  frmCustomCalendarControl.Hide
  
End Sub
Public Sub InitialiseAllDateToggleButtons()

  ' initialise all date toggle buttons so that they appear as not being pressed
  togDate01.Value = False
  togDate02.Value = False
  togDate03.Value = False
  togDate04.Value = False
  togDate05.Value = False
  togDate06.Value = False
  togDate07.Value = False
  togDate08.Value = False
  togDate09.Value = False
  togDate10.Value = False
  togDate11.Value = False
  togDate12.Value = False
  togDate13.Value = False
  togDate14.Value = False
  togDate15.Value = False
  togDate16.Value = False
  togDate17.Value = False
  togDate18.Value = False
  togDate19.Value = False
  togDate20.Value = False
  togDate21.Value = False
  togDate22.Value = False
  togDate23.Value = False
  togDate24.Value = False
  togDate25.Value = False
  togDate26.Value = False
  togDate27.Value = False
  togDate28.Value = False
  togDate29.Value = False
  togDate30.Value = False
  togDate31.Value = False
  togDate32.Value = False
  togDate33.Value = False
  togDate34.Value = False
  togDate35.Value = False
  togDate36.Value = False
  togDate37.Value = False
  togDate38.Value = False
  togDate39.Value = False
  togDate40.Value = False
  togDate41.Value = False
  togDate42.Value = False

  ' initialise the date label used for debugging
  lblDebugDate.Caption = ""

End Sub

Public Sub RedrawDayToggleButtons()

  Dim idxDayOffsetForStartOfMonth As Integer
  Dim idxDayOffset As Integer

  ' derive an offset used later to access the correct date toggle buttons
  Select Case Weekday(DateSerial(mSelectedYear, mSelectedMonth, 1))
         Case 1: idxDayOffsetForStartOfMonth = 0
         Case 2: idxDayOffsetForStartOfMonth = -1
         Case 3: idxDayOffsetForStartOfMonth = -2
         Case 4: idxDayOffsetForStartOfMonth = -3
         Case 5: idxDayOffsetForStartOfMonth = -4
         Case 6: idxDayOffsetForStartOfMonth = -5
         Case 7: idxDayOffsetForStartOfMonth = -6
  End Select
  
  ' format day toggle button 01
  idxDayOffset = 0
  With togDate01
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 02
  idxDayOffset = idxDayOffset + 1
  With togDate02
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 03
  idxDayOffset = idxDayOffset + 1
  With togDate03
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 04
  idxDayOffset = idxDayOffset + 1
  With togDate04
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 05
  idxDayOffset = idxDayOffset + 1
  With togDate05
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 06
  idxDayOffset = idxDayOffset + 1
  With togDate06
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 07
  idxDayOffset = idxDayOffset + 1
  With togDate07
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 08
  idxDayOffset = idxDayOffset + 1
  With togDate08
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 09
  idxDayOffset = idxDayOffset + 1
  With togDate09
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 10
  idxDayOffset = idxDayOffset + 1
  With togDate10
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 11
  idxDayOffset = idxDayOffset + 1
  With togDate11
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 12
  idxDayOffset = idxDayOffset + 1
  With togDate12
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 13
  idxDayOffset = idxDayOffset + 1
  With togDate13
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 14
  idxDayOffset = idxDayOffset + 1
  With togDate14
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 15
  idxDayOffset = idxDayOffset + 1
  With togDate15
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 16
  idxDayOffset = idxDayOffset + 1
  With togDate16
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 17
  idxDayOffset = idxDayOffset + 1
  With togDate17
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 18
  idxDayOffset = idxDayOffset + 1
  With togDate18
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 19
  idxDayOffset = idxDayOffset + 1
  With togDate19
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 20
  idxDayOffset = idxDayOffset + 1
  With togDate20
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 21
  idxDayOffset = idxDayOffset + 1
  With togDate21
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 22
  idxDayOffset = idxDayOffset + 1
  With togDate22
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 23
  idxDayOffset = idxDayOffset + 1
  With togDate23
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 24
  idxDayOffset = idxDayOffset + 1
  With togDate24
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 25
  idxDayOffset = idxDayOffset + 1
  With togDate25
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 26
  idxDayOffset = idxDayOffset + 1
  With togDate26
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 27
  idxDayOffset = idxDayOffset + 1
  With togDate27
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 28
  idxDayOffset = idxDayOffset + 1
  With togDate28
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 29
  idxDayOffset = idxDayOffset + 1
  With togDate29
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 30
  idxDayOffset = idxDayOffset + 1
  With togDate30
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 31
  idxDayOffset = idxDayOffset + 1
  With togDate31
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 32
  idxDayOffset = idxDayOffset + 1
  With togDate32
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 33
  idxDayOffset = idxDayOffset + 1
  With togDate33
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 34
  idxDayOffset = idxDayOffset + 1
  With togDate34
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 35
  idxDayOffset = idxDayOffset + 1
  With togDate35
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 36
  idxDayOffset = idxDayOffset + 1
  With togDate36
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 37
  idxDayOffset = idxDayOffset + 1
  With togDate37
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 38
  idxDayOffset = idxDayOffset + 1
  With togDate38
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 39
  idxDayOffset = idxDayOffset + 1
  With togDate39
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 40
  idxDayOffset = idxDayOffset + 1
  With togDate40
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 41
  idxDayOffset = idxDayOffset + 1
  With togDate41
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' format day toggle button 42
  idxDayOffset = idxDayOffset + 1
  With togDate42
      .Caption = Day(DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1)))
      .ControlTipText = DateAdd("d", idxDayOffset + idxDayOffsetForStartOfMonth, DateSerial(mSelectedYear, mSelectedMonth, 1))
      If MonthName(Month(.ControlTipText)) = cmbMonth.Text Then
         .ForeColor = RGB(0, 0, 128) ' day is in selected month therefore make font blue
         .Font.Bold = True
      Else
         .ForeColor = RGB(128, 128, 128) ' day is not in selected month therefore make font grey
         .Font.Bold = False
      End If
  End With

  ' initialise all toggle buttons so that they are not pressed
  Call InitialiseAllDateToggleButtons

  ' disable the OK button (it will be enabled again when the user clicks on a day button)
  cmdOK.Enabled = False
  cmdCancel.SetFocus

End Sub

Public Function UserIsADeveloper() As Boolean

  Dim CurrentUserName As String
  Dim ListOfDevelopers As String
  
  ListOfDevelopers = UCase(";kevin;kevin.clark;mickey mouse;bhavesh.patel;amit.rawat;paul.milner;" & _
                           "scott.spence;nigel,o'coffey;andrew.collings;niall.grant;")
                     
  CurrentUserName = ";" & UCase(Environ("USERNAME")) & ";"
  
  If InStr(1, ListOfDevelopers, CurrentUserName, vbTextCompare) > 0 Then
     UserIsADeveloper = True
  Else
     UserIsADeveloper = False
  End If
  
End Function

Public Sub SelectDefaultDay()

  ' select the correct toggle day button dependant upon the default day that has been passed into this form
  If Day(togDate01.ControlTipText) = mSelectedDay And Month(togDate01.ControlTipText) = mSelectedMonth Then togDate01.Value = True
  If Day(togDate02.ControlTipText) = mSelectedDay And Month(togDate02.ControlTipText) = mSelectedMonth Then togDate02.Value = True
  If Day(togDate03.ControlTipText) = mSelectedDay And Month(togDate03.ControlTipText) = mSelectedMonth Then togDate03.Value = True
  If Day(togDate04.ControlTipText) = mSelectedDay And Month(togDate04.ControlTipText) = mSelectedMonth Then togDate04.Value = True
  If Day(togDate05.ControlTipText) = mSelectedDay And Month(togDate05.ControlTipText) = mSelectedMonth Then togDate05.Value = True
  If Day(togDate06.ControlTipText) = mSelectedDay And Month(togDate06.ControlTipText) = mSelectedMonth Then togDate06.Value = True
  If Day(togDate07.ControlTipText) = mSelectedDay And Month(togDate07.ControlTipText) = mSelectedMonth Then togDate07.Value = True
  If Day(togDate08.ControlTipText) = mSelectedDay And Month(togDate08.ControlTipText) = mSelectedMonth Then togDate08.Value = True
  If Day(togDate09.ControlTipText) = mSelectedDay And Month(togDate09.ControlTipText) = mSelectedMonth Then togDate09.Value = True
  If Day(togDate10.ControlTipText) = mSelectedDay And Month(togDate10.ControlTipText) = mSelectedMonth Then togDate10.Value = True
  If Day(togDate11.ControlTipText) = mSelectedDay And Month(togDate11.ControlTipText) = mSelectedMonth Then togDate11.Value = True
  If Day(togDate12.ControlTipText) = mSelectedDay And Month(togDate12.ControlTipText) = mSelectedMonth Then togDate12.Value = True
  If Day(togDate13.ControlTipText) = mSelectedDay And Month(togDate13.ControlTipText) = mSelectedMonth Then togDate13.Value = True
  If Day(togDate14.ControlTipText) = mSelectedDay And Month(togDate14.ControlTipText) = mSelectedMonth Then togDate14.Value = True
  If Day(togDate15.ControlTipText) = mSelectedDay And Month(togDate15.ControlTipText) = mSelectedMonth Then togDate15.Value = True
  If Day(togDate16.ControlTipText) = mSelectedDay And Month(togDate16.ControlTipText) = mSelectedMonth Then togDate16.Value = True
  If Day(togDate17.ControlTipText) = mSelectedDay And Month(togDate17.ControlTipText) = mSelectedMonth Then togDate17.Value = True
  If Day(togDate18.ControlTipText) = mSelectedDay And Month(togDate18.ControlTipText) = mSelectedMonth Then togDate18.Value = True
  If Day(togDate19.ControlTipText) = mSelectedDay And Month(togDate19.ControlTipText) = mSelectedMonth Then togDate19.Value = True
  If Day(togDate20.ControlTipText) = mSelectedDay And Month(togDate20.ControlTipText) = mSelectedMonth Then togDate20.Value = True
  If Day(togDate21.ControlTipText) = mSelectedDay And Month(togDate21.ControlTipText) = mSelectedMonth Then togDate21.Value = True
  If Day(togDate22.ControlTipText) = mSelectedDay And Month(togDate22.ControlTipText) = mSelectedMonth Then togDate22.Value = True
  If Day(togDate23.ControlTipText) = mSelectedDay And Month(togDate23.ControlTipText) = mSelectedMonth Then togDate23.Value = True
  If Day(togDate24.ControlTipText) = mSelectedDay And Month(togDate24.ControlTipText) = mSelectedMonth Then togDate24.Value = True
  If Day(togDate25.ControlTipText) = mSelectedDay And Month(togDate25.ControlTipText) = mSelectedMonth Then togDate25.Value = True
  If Day(togDate26.ControlTipText) = mSelectedDay And Month(togDate26.ControlTipText) = mSelectedMonth Then togDate26.Value = True
  If Day(togDate27.ControlTipText) = mSelectedDay And Month(togDate27.ControlTipText) = mSelectedMonth Then togDate27.Value = True
  If Day(togDate28.ControlTipText) = mSelectedDay And Month(togDate28.ControlTipText) = mSelectedMonth Then togDate28.Value = True
  If Day(togDate29.ControlTipText) = mSelectedDay And Month(togDate29.ControlTipText) = mSelectedMonth Then togDate29.Value = True
  If Day(togDate30.ControlTipText) = mSelectedDay And Month(togDate30.ControlTipText) = mSelectedMonth Then togDate30.Value = True
  If Day(togDate31.ControlTipText) = mSelectedDay And Month(togDate31.ControlTipText) = mSelectedMonth Then togDate31.Value = True
  If Day(togDate32.ControlTipText) = mSelectedDay And Month(togDate32.ControlTipText) = mSelectedMonth Then togDate32.Value = True
  If Day(togDate33.ControlTipText) = mSelectedDay And Month(togDate33.ControlTipText) = mSelectedMonth Then togDate33.Value = True
  If Day(togDate34.ControlTipText) = mSelectedDay And Month(togDate34.ControlTipText) = mSelectedMonth Then togDate34.Value = True
  If Day(togDate35.ControlTipText) = mSelectedDay And Month(togDate35.ControlTipText) = mSelectedMonth Then togDate35.Value = True
  If Day(togDate36.ControlTipText) = mSelectedDay And Month(togDate36.ControlTipText) = mSelectedMonth Then togDate36.Value = True
  If Day(togDate37.ControlTipText) = mSelectedDay And Month(togDate37.ControlTipText) = mSelectedMonth Then togDate37.Value = True
  If Day(togDate38.ControlTipText) = mSelectedDay And Month(togDate38.ControlTipText) = mSelectedMonth Then togDate38.Value = True
  If Day(togDate39.ControlTipText) = mSelectedDay And Month(togDate39.ControlTipText) = mSelectedMonth Then togDate39.Value = True
  If Day(togDate40.ControlTipText) = mSelectedDay And Month(togDate40.ControlTipText) = mSelectedMonth Then togDate40.Value = True
  If Day(togDate41.ControlTipText) = mSelectedDay And Month(togDate41.ControlTipText) = mSelectedMonth Then togDate41.Value = True
  If Day(togDate42.ControlTipText) = mSelectedDay And Month(togDate42.ControlTipText) = mSelectedMonth Then togDate42.Value = True

End Sub

Private Sub togDate01_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate01
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate02_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate02
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub

Private Sub togDate03_Click()
  
  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate03
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate04_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate04
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate05_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate05
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate06_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate06
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate07_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate07
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate08_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate08
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate09_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate09
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate10_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate10
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate11_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate11
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate12_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate12
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate13_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate13
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate14_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate14
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate15_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate15
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate16_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate16
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate17_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate17
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate18_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate18
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate19_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate19
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate20_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate20
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate21_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate21
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate22_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate22
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate23_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate23
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate24_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate24
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate25_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate25
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate26_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate26
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate27_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate27
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate28_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate28
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate29_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate29
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate30_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate30
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate31_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate31
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate32_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate32
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate33_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate33
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate34_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate34
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate35_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate35
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate36_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate36
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate37_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate37
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate38_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate38
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate39_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate39
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate40_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate40
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate41_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate41
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub
Private Sub togDate42_Click()

  ' if event handling has been switched off then exit this event handler
  If IgnoreEvents = True Then
     Exit Sub
  End If

  ' change the state of the day toggle button as appropriate
  With togDate42
       If .Value = True Then
          IgnoreEvents = True
          Call InitialiseAllDateToggleButtons
          .Value = True
          mSelectedDay = Day(.ControlTipText)
          mSelectedMonth = Month(.ControlTipText)
          mSelectedYear = Year(.ControlTipText)
          lblDebugDate.Caption = .ControlTipText
          cmdOK.Enabled = True
          IgnoreEvents = False
       Else
          lblDebugDate.Caption = ""
          cmdOK.Enabled = False
       End If
  End With

End Sub

