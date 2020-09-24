VERSION 5.00
Begin VB.Form frmSQS 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "SuperQickSort"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   Icon            =   "frmSQSBin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ComboBox cboSize 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmSQSBin.frx":000C
      Left            =   2550
      List            =   "frmSQSBin.frx":002B
      Style           =   2  'Dropdown-Liste
      TabIndex        =   8
      Top             =   75
      Width           =   825
   End
   Begin VB.ComboBox cboVol 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmSQSBin.frx":005C
      Left            =   495
      List            =   "frmSQSBin.frx":0075
      Style           =   2  'Dropdown-Liste
      TabIndex        =   7
      Top             =   75
      Width           =   1185
   End
   Begin VB.CommandButton btSort1 
      Caption         =   "Sort"
      Height          =   495
      Left            =   1733
      TabIndex        =   0
      Top             =   2310
      Width           =   1215
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      Caption         =   "Ratio"
      Height          =   195
      Index           =   3
      Left            =   270
      TabIndex        =   11
      Top             =   1605
      Width           =   375
   End
   Begin VB.Label lbRatio 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2700
      TabIndex        =   10
      Top             =   1605
      Width           =   75
   End
   Begin VB.Label lbTotalVol 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   255
      TabIndex        =   9
      Top             =   585
      Width           =   75
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Sort                     Strings of        18    Chars average"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   6
      Top             =   135
      Width           =   4590
   End
   Begin VB.Label lbOK 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2295
      TabIndex        =   5
      Top             =   1845
      Width           =   105
   End
   Begin VB.Label lbSQ 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2700
      TabIndex        =   4
      Top             =   1320
      Width           =   75
   End
   Begin VB.Label lbCQ 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2700
      TabIndex        =   3
      Top             =   1095
      Width           =   75
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      Caption         =   "Super Quickie"
      Height          =   195
      Index           =   1
      Left            =   270
      TabIndex        =   2
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      Caption         =   "Conventional Quicksort"
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   1
      Top             =   1095
      Width           =   1650
   End
End
Attribute VB_Name = "frmSQS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private PerfFreq        As Currency 'hi speed timer
Private PerfCnt1        As Currency
Private PerfCnt2        As Currency

Private ElemsB()        As String
Private ElemsA()        As String 'two identical tables
Private AvgLen          As Long
Private cConventional   As clsConventional
Private cQuickie        As clsQuickie

Private Sub btSort1_Click()

  Dim TotalVol  As Long
  Dim Timing1   As Currency
  Dim Timing2   As Currency
  Dim i         As Long

    Enabled = False
    lbSQ = vbNullString
    lbCQ = vbNullString
    lbOK = vbNullString
    lbRatio = vbNullString
    lbTotalVol = vbNullString

    If UBound(ElemsA) = 1000000 Then
        For i = 1 To UBound(ElemsA)
            ElemsA(i) = CStr(Rnd) & CStr(Rnd)
            ElemsB(i) = ElemsA(i)
            TotalVol = TotalVol + Len(ElemsA(i))
        Next i
      Else 'NOT UBOUND(ELEMSA)...
        For i = 1 To UBound(ElemsA)
            ElemsA(i) = Left$(CStr(Rnd) & String$(2 * AvgLen, "a"), (Rnd + 0.5) * AvgLen)
            ElemsB(i) = ElemsA(i)
            TotalVol = TotalVol + Len(ElemsA(i))
        Next i
    End If
    lbTotalVol = "Total Volume is " & Format$(TotalVol, "#,0") & " Unicode Chars"
    DoEvents

    'time conventional sort
    QueryPerformanceCounter PerfCnt1
    cConventional.QuicksortConven ElemsA, 1, UBound(ElemsA)
    QueryPerformanceCounter PerfCnt2
    Timing1 = (PerfCnt2 - PerfCnt1) / PerfFreq * 1000
    lbCQ = Format$(Timing1, "#0.00 mSec ")
    DoEvents

    'time superquickie1
    QueryPerformanceCounter PerfCnt1
    cQuickie.SuperQuickie ElemsB, 1, UBound(ElemsB)
    QueryPerformanceCounter PerfCnt2
    Timing2 = (PerfCnt2 - PerfCnt1) / PerfFreq * 1000
    lbSQ = Format$(Timing2, "#0.00 mSec ")
    DoEvents

    lbRatio = Format$(Timing1 / Timing2, "#0.0")
    DoEvents

    'check
    For i = 1 To UBound(ElemsA)
        If ElemsA(i) <> ElemsB(i) Or StrComp(ElemsA(i), ElemsA(i - 1)) < 0 Then
            Exit For 'loop varying i
        End If
    Next i
    If i > UBound(ElemsA) Then
        lbOK = "Sorts okay"
      Else 'NOT I...
        lbOK = "Sorts failed"
    End If
    Enabled = True

End Sub

Private Sub cboSize_Click()

    AvgLen = cboSize.List(cboSize.ListIndex)

End Sub

Private Sub cboVol_Click()

    Enabled = False
    cboSize.Visible = (cboVol.List(cboVol.ListIndex) <> 1000000)
    DoEvents
    ReDim ElemsA(0 To cboVol.List(cboVol.ListIndex))
    ReDim ElemsB(0 To UBound(ElemsA))
    ElemsA(0) = Chr$(0) 'not sorted, just for scanning
    ElemsB(0) = Chr$(0) 'not sorted, just for scanning
    Enabled = True

End Sub

Private Sub Form_Load()

    Set cConventional = New clsConventional
    Set cQuickie = New clsQuickie
    QueryPerformanceFrequency PerfFreq
    cboVol.ListIndex = 3 '1000 strings
    cboSize.ListIndex = 7 '2000 chars

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Hide
    DoEvents

End Sub

':) Ulli's VB Code Formatter V2.23.12 (2007-Mrz-25 13:48)  Decl: 13  Code: 103  Total: 116 Lines
':) CommentOnly: 5 (4,3%)  Commented: 8 (6,9%)  Empty: 24 (20,7%)  Max Logic Depth: 3
