VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Internet History,cookies viewer by Mahmoud Faraj"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   11490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Delete selected items"
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Delete all history items"
      Height          =   495
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "History"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Temp.files"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Cookies"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.ComboBox cmbday 
      Height          =   315
      Left            =   3720
      TabIndex        =   2
      Text            =   "choose date"
      Top             =   360
      Width           =   1695
   End
   Begin SHDocVwCtl.WebBrowser w1 
      Height          =   3375
      Left            =   0
      TabIndex        =   1
      Top             =   3480
      Width           =   11295
      ExtentX         =   19923
      ExtentY         =   5953
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "url"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "lastaccesstime"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "expired time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "hitrate"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Items number"
      Height          =   255
      Left            =   6840
      TabIndex        =   9
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Select date"
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label nlabel 
      Caption         =   " "
      Height          =   255
      Left            =   7080
      TabIndex        =   7
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   " "
      Height          =   255
      Left            =   5760
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub getcachentry(sdate As Date)
Dim xdate As Date
nlabel.Caption = ""
ListView1.ListItems.Clear
Dim URL() As Internet_Cache_Entry
Dim URLHistory() As Internet_Cache_Entry
Dim Cookies() As Internet_Cache_Entry
x = GetURLCache(URL(), URLHistory(), Cookies())
If Option1.Value = True Then
For N = 1 To UBound(URLHistory)
x = InStr(URLHistory(N).SourceUrlName, "@")
xurl = Right(URLHistory(N).SourceUrlName, Len(URLHistory(N).SourceUrlName) - x)
If x > 0 Then
xcontent = Mid(xurl, x, 23)
End If
xdate = DateValue(URLHistory(N).LastAccessTime)
If xdate = sdate And Left$(xurl, 4) = "http" And Right(xurl, 3) <> "gif" And Right(xurl, 3) <> "jpg" And Right(xurl, 3) <> "zip" Then
 
i = i + 1

ListView1.ListItems.Add , "s" & i, URLHistory(N).SourceUrlName
k = k + 1
ListView1.ListItems(i).ListSubItems.Add , "m" & k & i, URLHistory(N).LastAccessTime
k = k + 1
ListView1.ListItems(i).ListSubItems.Add , "m" & k & i, URLHistory(N).ExpireTime
k = k + 1
ListView1.ListItems(i).ListSubItems.Add , "m" & k & i, URLHistory(N).HitRate



End If
Next N
ElseIf Option2.Value = True Then
For N = 1 To UBound(URL)
''x = InStr(URLHistory(N).SourceUrlName, "@")
xurl = URL(N).SourceUrlName
xdate = DateValue(URL(N).LastAccessTime)
If xdate = sdate Then


i = i + 1

ListView1.ListItems.Add , "s" & i, URL(N).SourceUrlName
k = k + 1
ListView1.ListItems(i).ListSubItems.Add , "m" & k & i, URL(N).LastAccessTime
k = k + 1
ListView1.ListItems(i).ListSubItems.Add , "m" & k & i, URL(N).ExpireTime
k = k + 1
ListView1.ListItems(i).ListSubItems.Add , "m" & k & i, URL(N).HitRate

End If
Next N
ElseIf Option3.Value = True Then
For N = 1 To UBound(Cookies)
''x = InStr(URLHistory(N).SourceUrlName, "@")
xurl = Cookies(N).LocalFileName
xdate = DateValue(Cookies(N).LastAccessTime)
If xdate = sdate Then

i = i + 1

ListView1.ListItems.Add , "s" & i, Cookies(N).SourceUrlName
ListView1.ListItems.Item("s" & i).Tag = Cookies(N).LocalFileName
k = k + 1
ListView1.ListItems(i).ListSubItems.Add , "m" & k & i, Cookies(N).LastAccessTime
k = k + 1
ListView1.ListItems(i).ListSubItems.Add , "m" & k & i, Cookies(N).ExpireTime
k = k + 1
ListView1.ListItems(i).ListSubItems.Add , "m" & k & i, Cookies(N).HitRate


End If
Next N
End If
nlabel.Caption = ListView1.ListItems.Count
End Sub
Public Sub fillday2()
On Error GoTo rt

Dim sdate As Date

sdate = DateAdd("d", -1, Date)
For i = 0 To 30
sdate = DateAdd("d", -i, Date)
cmbday.AddItem sdate
Next i
cmbday.ListIndex = 0
sdate = DateValue(cmbday.Text)
''getcachentry sdate


Exit Sub
rt:
MsgBox Error$
Resume rte:
rte:

End Sub

Private Sub cmbday_Change()
fdate = cmbday.Text
xk = Weekday(fdate, vbSunday)
Select Case xk
Case 1
mtdate = "Monday"
Case 2
mtdate = "Sunday"
Case 3
mtdate = "Tuesday"
Case 4
mtdate = "Wenesday"
Case 5
mtdate = "Thursday"
Case 6
mtdate = "Friday"
Case 7
mtdate = "Saturday"
End Select
Label1.Caption = mtdate

End Sub

Private Sub cmbday_Click()
Dim sdate As Date
fdate = cmbday.Text
xk = Weekday(fdate, vbSunday)
Select Case xk
Case 1
mtdate = "Monday"
Case 2
mtdate = "Sunday"
Case 3
mtdate = "Tuesday"
Case 4
mtdate = "Wenesday"
Case 5
mtdate = "Thursday"
Case 6
mtdate = "Friday"
Case 7
mtdate = "Saturday"
End Select
Label1.Caption = mtdate
sdate = DateValue(cmbday.Text)
getcachentry sdate
End Sub

Private Sub Command1_Click()
Dim answer%
Dim xdone As Boolean
Dim sdate As Date
Dim liste() As Internet_Cache_Entry
answer = MsgBox("All internet history items will be deleted", vbYesNo, "Warning")
If answer = 6 Then
xdone = DeleteUrlCache(liste)
If xdone = True Then
MsgBox "all Item are deleted"
sdate = DateValue(cmbday.Text)
getcachentry sdate
End If
End If
End Sub

Private Sub Command2_Click()
Dim answer%
Dim selecteditem As String
Dim sdate As Date
Dim xdone As Boolean

selecteditem = ListView1.selecteditem.Text
If selecteditem = "" Then Exit Sub
answer = MsgBox("Selected internet history item will be deleted", vbYesNo, "Warning")
If answer = 6 Then
xdone = deleteselecteditem(selecteditem)
If xdone = True Then
MsgBox "Item is delected"
sdate = DateValue(cmbday.Text)
getcachentry sdate
ListView1.SetFocus
End If
End If

End Sub

Private Sub Form_Load()
fillday2
Option1.Value = True
fdate = cmbday.Text
xk = Weekday(fdate, vbSunday)
Select Case xk
Case 1
mtdate = "Monday"
Case 2
mtdate = "Sunday"
Case 3
mtdate = "Tuesday"
Case 4
mtdate = "Wenesday"
Case 5
mtdate = "Thursday"
Case 6
mtdate = "Friday"
Case 7
mtdate = "Saturday"
End Select
Label1.Caption = mtdate
w1.Offline = True
End Sub

Private Sub ListView1_Click()
w1.Offline = True
Dim xurl$
xurl = ListView1.selecteditem.Text
If xurl = "" Then Exit Sub
If Option1.Value = True Then

x = InStr(xurl, "@")
xurl = Right(xurl, Len(xurl) - x)
w1.Navigate xurl
ElseIf Option2.Value = True Then
w1.Navigate xurl
ElseIf Option3.Value = True Then
xurl = ListView1.selecteditem.Tag
w1.Navigate xurl
End If

End Sub

Private Sub Option1_Click()
Dim sdate As Date
If Option1.Value = True Then
sdate = DateValue(cmbday.Text)
getcachentry sdate
End If
End Sub

Private Sub Option2_Click()
Dim sdate As Date
If Option2.Value = True Then
sdate = DateValue(cmbday.Text)
getcachentry sdate
End If
End Sub

Private Sub Option3_Click()
Dim sdate As Date
If Option3.Value = True Then
sdate = DateValue(cmbday.Text)
getcachentry sdate
End If
End Sub
