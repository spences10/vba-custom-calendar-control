VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTestHarness 
   Caption         =   "Test Harness"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3480
   OleObjectBlob   =   "frmTestHarness.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmTestHarness"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQuit_Click()

  frmTestHarness.Hide
  Unload frmTestHarness

End Sub

Private Sub imgCalendarButton_Click()

  If txtDate.Text = "" Then
     '// do nothing
  Else
     If IsDate(txtDate.Text) = True Then
        Load frmCustomCalendarControl
        frmCustomCalendarControl.SelectedDayNumber = Day(txtDate.Text)
        frmCustomCalendarControl.SelectedMonthNumber = Month(txtDate.Text)
        frmCustomCalendarControl.SelectedYearNumber = Year(txtDate.Text)
     End If
  End If

  frmCustomCalendarControl.Show
  
  If frmCustomCalendarControl.SelectedDayNumber = 0 And _
     frmCustomCalendarControl.SelectedMonthNumber = 0 And _
     frmCustomCalendarControl.SelectedYearNumber = 0 Then
     '// user click on the cancel button in the calendar control therefore do nothing
  Else
     txtDate.Text = DateSerial(frmCustomCalendarControl.SelectedYearNumber, _
                               frmCustomCalendarControl.SelectedMonthNumber, _
                               frmCustomCalendarControl.SelectedDayNumber)
     '// the following properties are also available from the customer control if you need them
     Debug.Print "frmCustomCalendarControl.SelectedDateDDMMYYYY = " & frmCustomCalendarControl.SelectedDateDDMMYYYY
     Debug.Print "frmCustomCalendarControl.SelectedDayString = " & frmCustomCalendarControl.SelectedDayString
     Debug.Print "frmCustomCalendarControl.SelectedMonthString = " & frmCustomCalendarControl.SelectedMonthString
     Debug.Print "frmCustomCalendarControl.SelectedYearString = " & frmCustomCalendarControl.SelectedYearString
  End If

  Unload frmCustomCalendarControl
  
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

  Cancel = True
  
End Sub
