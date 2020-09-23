VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "s&elect row..."
      Height          =   285
      Left            =   225
      TabIndex        =   6
      Top             =   990
      Width           =   1590
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&set the cell text for..."
      Height          =   285
      Left            =   225
      TabIndex        =   5
      Top             =   585
      Width           =   1635
   End
   Begin VB.ComboBox comboColumn 
      Height          =   315
      Left            =   3060
      TabIndex        =   4
      Text            =   "column"
      Top             =   180
      Width           =   915
   End
   Begin VB.ComboBox comboRow 
      Height          =   315
      Left            =   1980
      TabIndex        =   3
      Text            =   "row"
      Top             =   180
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2025
      TabIndex        =   2
      Top             =   585
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "get the cell text for.."
      Height          =   285
      Left            =   225
      TabIndex        =   1
      Top             =   180
      Width           =   1590
   End
   Begin projListGridControl.ListGrid ListGrid1 
      Height          =   1545
      Left            =   270
      TabIndex        =   0
      Top             =   1395
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   2725
      Header          =   " stock symbol  | stock type  |  stock scope | current price "
      header_style    =   0
      header_text_color=   16711680
      BeginProperty header_font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty list_font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      list_forecolor  =   16711680
      list_backcolor  =   16777215
      list_borderstyle=   0
      allowsort_onheader_click=   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Change()
 
End Sub

Private Sub Command1_Click()

 MsgBox ListGrid1.cell_text(comboRow.ListIndex, comboColumn.ListIndex)

End Sub

Private Sub Command2_Click()
 
  ListGrid1.cell_text(comboRow.ListIndex, comboColumn.ListIndex) = Text1
 
End Sub

Private Sub Command3_Click()
  
 ListGrid1.selected_row = comboRow.ListIndex

End Sub

Private Sub Form_Load()
  
  Dim lcnt&
  
  With ListGrid1
    .horizontal_scrollbar = True
    
    .tab_points 47, 86, 128
    .add_row "vivax|mutual fund|U.S.|$21.20"
    .add_row "vtsmx|mutual fund|U.S.|$28.50"
    .add_row "coke|stock|worldwide|check back later gator"
    .add_row "siri|stock|U.S.|$5.60"
    .add_row "yhoo|stock|overseas|10.22"
    .add_row "ggle|coins|Europe|110.22"
  End With
  
  For lcnt = 0 To 3
    If lcnt < 3 Then comboRow.AddItem lcnt
    comboColumn.AddItem lcnt
  Next lcnt
  
End Sub

