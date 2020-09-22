VERSION 5.00
Begin VB.Form frmMst 
   AutoRedraw      =   -1  'True
   Caption         =   "Minimum Spanning Tree Demo"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdKruskalMST 
      Caption         =   "S&how Kruskal's MST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   6840
      Width           =   6615
   End
   Begin VB.CommandButton cmdPrimMST 
      Caption         =   "&Show Prim's MST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   6615
   End
   Begin VB.OptionButton optKruskal 
      Caption         =   "&K r u s k a l"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.OptionButton optPrim 
      Caption         =   "&P r i m"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   4560
      Width           =   6615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1560
      Width           =   6615
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run it"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim m() As Double
Dim mstP() As Double
Dim mstK() As Double

Private Sub cmdKruskalMST_Click()
    Dim a As Integer, b As Integer
    On Error Resume Next
    
    frmMatrix.txtMst.Text = ""
    For a = 1 To UBound(mstK)
        For b = 1 To UBound(mstK)
            frmMatrix.txtMst.Text = frmMatrix.txtMst.Text & mstK(a, b) & vbTab
        Next
        frmMatrix.txtMst.Text = frmMatrix.txtMst.Text & vbCrLf
    Next
    
    frmMatrix.Show
End Sub

Private Sub cmdPrimMST_Click()
    Dim a As Integer, b As Integer
    On Error Resume Next
    
    frmMatrix.txtMst.Text = ""
    For a = 1 To UBound(mstP)
        For b = 1 To UBound(mstP)
            frmMatrix.txtMst.Text = frmMatrix.txtMst.Text & mstP(a, b) & vbTab
        Next
        frmMatrix.txtMst.Text = frmMatrix.txtMst.Text & vbCrLf
    Next
    
    frmMatrix.Show
End Sub

Private Sub cmdRun_Click()
    Dim initText As String
    ReDim m(7, 7) As Double
    
    m(1, 2) = 7
    m(1, 4) = 5
    m(2, 3) = 8
    m(2, 4) = 9
    m(2, 5) = 7
    m(3, 5) = 5
    m(4, 5) = 15
    m(4, 6) = 6
    m(5, 6) = 8
    m(5, 7) = 9
    m(6, 7) = 11
    '------------
    m(2, 1) = 7
    m(4, 1) = 5
    m(3, 2) = 8
    m(4, 2) = 9
    m(5, 2) = 7
    m(5, 3) = 5
    m(5, 4) = 15
    m(6, 4) = 6
    m(6, 5) = 8
    m(7, 5) = 9
    m(7, 6) = 11
    
    initText = "Number of edges from the initial matrix = " & numberOfEdge(m) & _
    vbCrLf & "Total weight of the initial matrix = " & countTotalWeight(m)
    
    If optPrim.Value Then
        Text1.Text = initText
        ReDim mstP(7, 7) As Double
        Call prim(m, 1, mstP)
        Text1.Text = Text1.Text & vbCrLf & "Number of edges from the MST = " & numberOfEdge(mstP)
        Text1.Text = Text1.Text & vbCrLf & "Total weight of the MST = " & countTotalWeight(mstP)
    Else
        Text2.Text = initText
        ReDim mstK(7, 7) As Double
        Call kruskal(m, mstK)
        Text2.Text = Text2.Text & vbCrLf & "Number of edges from the MST = " & numberOfEdge(mstK)
        Text2.Text = Text2.Text & vbCrLf & "Total weight of the MST = " & countTotalWeight(mstK)
    End If
End Sub
