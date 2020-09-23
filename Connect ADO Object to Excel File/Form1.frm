VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect ADO Object to Excel File"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6800
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1037
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1037
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get range of data from Excel Sheet"
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get all DATA from Excel Sheet"
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   4200
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Option Explicit

Private Sub Command1_Click()
    
    Set rs = New ADODB.Recordset
    '--- open recordset
    '--- "Members" is the name of one of the sheets
    '--- at the Excel file
    
    rs.Open "SELECT * FROM [Members$] ", cn, adOpenDynamic, adLockOptimistic
    
    Set DataGrid1.DataSource = rs
    
    
End Sub



Private Sub Command2_Click()

    Set rs = New ADODB.Recordset
    
    '--- this time select a range of cells from the XLS file
    '--- notice that Salary is one of the sheets at the Excel File
    rs.Open "SELECT * FROM [Salary$A1:B2] ", cn, adOpenDynamic, adLockOptimistic
    
    Set DataGrid1.DataSource = rs
    
End Sub



Private Sub Form_Load()

On Error GoTo ErrHandler
    Set cn = New ADODB.Connection
    
    ' -- connection provider
    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    
    '--- create the connection to Excel File
    '--- notice for the "Extended Property"
    '--- for Excel 97/2000/2002 use Excel 8.0
    '--- for Excel 95 use Excel 5.0
    cn.ConnectionString = _
        "Data Source= " & App.Path & "/Book1.xls;" & _
        "Extended Properties=Excel 8.0;"
    cn.CursorLocation = adUseClient
    cn.Open
    
Exit Sub
ErrHandler:
    MsgBox "Can not establish the connection"
End Sub

Private Sub Command3_Click()
    MsgBox "Example by Gil Shabthai 10/09/2002", vbInformation, ""
    End
End Sub
