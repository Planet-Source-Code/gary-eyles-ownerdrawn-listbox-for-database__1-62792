VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check11 
      Caption         =   "Column Color"
      Height          =   255
      Left            =   6000
      TabIndex        =   19
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CheckBox Check10 
      Caption         =   "Vertical Header"
      Height          =   255
      Left            =   6000
      TabIndex        =   18
      Top             =   1200
      Width           =   1575
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   495
      Left            =   7200
      Max             =   100
      Min             =   22
      TabIndex        =   17
      Top             =   240
      Value           =   22
      Width           =   255
   End
   Begin VB.CheckBox Check9 
      Caption         =   "xp Theme"
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   1560
      Width           =   1455
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   495
      Left            =   6360
      Max             =   100
      Min             =   5
      TabIndex        =   15
      Top             =   240
      Value           =   20
      Width           =   255
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Column Fixed Width"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Jump to"
      Height          =   375
      Left            =   1200
      TabIndex        =   13
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Text            =   "2000"
      Top             =   960
      Width           =   975
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Enabled"
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   840
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Multiselect"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Foucs Rect"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Picture"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   240
      Picture         =   "Datalist.frx":0000
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Header"
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   840
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Grid Lines"
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLICK ME FIRST"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin Project1.Datalist Datalist1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9975
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Headercolor     =   -2147483626
      HeaderTextColor =   192
      ScrollbarBARcolor=   12632256
      BackColor       =   16777215
      GridlineColor   =   12632256
      XPtheme         =   0   'False
      IsEditable      =   -1  'True
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "HOTEL.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PostcodesXY"
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "20"
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    Datalist1.ShowGrid = Check1.Value
End Sub

Private Sub Check10_Click()
    Datalist1.VerticalHeader = Check10.Value
End Sub

Private Sub Check11_Click()
    Datalist1.UseColumnColor = Check11.Value
End Sub

Private Sub Check2_Click()
    Datalist1.ShowHeader = Check2.Value
End Sub

Private Sub Check3_Click()
    Datalist1.ShowCheck = Check3.Value
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
    Set Datalist1.Picture = Picture1.Image
Else
    Set Datalist1.Picture = Nothing
End If
End Sub

Private Sub Check5_Click()
    Datalist1.ShowFocusRect = Check5.Value
End Sub

Private Sub Check6_Click()
    Datalist1.MultiSelect = Check6.Value
End Sub

Private Sub Check7_Click()
    Datalist1.Enabled = Check7.Value
End Sub

Private Sub Check8_Click()
    Datalist1.ColumnFixedWidth = Check8.Value
End Sub

Private Sub Check9_Click()
    Datalist1.XPtheme = Not Datalist1.XPtheme
End Sub

Private Sub Command1_Click()
    Set Datalist1.theData = Data1
    Datalist1.ShowHeader = True
    Datalist1.ShowCheck = True
    
    Dim c As Long
    For c = 0 To Datalist1.ColumnCount - 1
        Datalist1.SetColColor c, QBColor(Int(Rnd * 14) + 1)
        'Debug.Print c
    Next

    Datalist1.DrawList

End Sub

Private Sub Command2_Click()
    Datalist1.Selected = Val(Text1.Text)
End Sub

Private Sub Datalist1_HeaderClick(colIndex As Variant)
    'Data1.RecordSource = Datalist1.SortSql(colIndex)
    'Data1.Refresh
    'Data1.Recordset.Movelast
    'Datalist1.DrawList
End Sub

Private Sub Datalist1_ItemSelected(tIndex As Long)
If tIndex > -1 Then
    Data1.Recordset.AbsolutePosition = tIndex
    Label1.Caption = Data1.Recordset.Fields(0)
End If
End Sub

Private Sub Form_Activate()
'Caption = Datalist1.hwnd
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Datalist1.Width = ScaleWidth - Datalist1.left * 2
    Datalist1.Height = ScaleHeight - Datalist1.tOp - Datalist1.left
End Sub

Private Sub VScroll1_Change()
    Datalist1.RowHeight = VScroll1.Value
    Label2.Caption = VScroll1.Value
End Sub

Private Sub VScroll2_Change()
    Datalist1.HeaderHeight = VScroll2.Value
End Sub
