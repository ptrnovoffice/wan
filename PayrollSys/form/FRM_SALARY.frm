VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form FRM_SALARY 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "List Salary Employe"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   10620
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControlPage TabBack 
      Height          =   7935
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11415
      _Version        =   851970
      _ExtentX        =   20135
      _ExtentY        =   13996
      _StockProps     =   1
      BackColor       =   4772349
      PictureAlignment=   32
      Begin XtremeSuiteControls.PushButton cmdList 
         Height          =   735
         Index           =   0
         Left            =   8880
         TabIndex        =   1
         Top             =   7080
         Visible         =   0   'False
         Width           =   735
         _Version        =   851970
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Hapus"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton cmdList 
         Height          =   735
         Index           =   1
         Left            =   8040
         TabIndex        =   2
         Top             =   7080
         Visible         =   0   'False
         Width           =   735
         _Version        =   851970
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Add"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton cmdList 
         Height          =   735
         Index           =   2
         Left            =   9720
         TabIndex        =   3
         Top             =   7080
         Width           =   735
         _Version        =   851970
         _ExtentX        =   1296
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Edit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         ImageAlignment  =   6
         TextImageRelation=   1
         ImageGap        =   0
      End
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid 
         Height          =   6255
         Left            =   120
         OleObjectBlob   =   "FRM_SALARY.frx":0000
         TabIndex        =   0
         Top             =   720
         Width           =   10335
      End
      Begin XtremeSuiteControls.ComboBox cmbFilter 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Barang Type "
         Top             =   360
         Width           =   2055
         _Version        =   851970
         _ExtentX        =   3625
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Sorted          =   -1  'True
         Appearance      =   1
         UseVisualStyle  =   0   'False
         AutoComplete    =   -1  'True
         EnableMarkup    =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbFilter 
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Barang Satuan"
         Top             =   360
         Width           =   1695
         _Version        =   851970
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Sorted          =   -1  'True
         Appearance      =   1
         UseVisualStyle  =   0   'False
         AutoComplete    =   -1  'True
         EnableMarkup    =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cmbFilter 
         Height          =   315
         Index           =   2
         Left            =   3840
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Barang Kategori"
         Top             =   360
         Width           =   1695
         _Version        =   851970
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Sorted          =   -1  'True
         Appearance      =   1
         UseVisualStyle  =   0   'False
         AutoComplete    =   -1  'True
         EnableMarkup    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblVar 
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   10
         Top             =   120
         Width           =   1095
         _Version        =   851970
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Golongan"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblVar 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   855
         _Version        =   851970
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cabang"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblVar 
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   8
         Top             =   120
         Width           =   1215
         _Version        =   851970
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Departemen"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Rounded MT Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
   End
End
Attribute VB_Name = "FRM_SALARY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CLM1 As Variant
Dim lR As Long
Dim i As Integer

Private Sub cmbFilter_Click(Index As Integer)
'=========================
'---COMBO FILTER CLICK ---
'=========================
If Index = 0 Or Index = 1 Or Index = 2 Or Index = 3 Or Index = 4 Then
    GridFill
End If
End Sub

Private Sub cmdList_Click(Index As Integer)
Dim defId As String
On Error GoTo ErrorLabel
LOADING.Show
LOADING.SetParm Me, 10


Select Case Index
    Case 0 'tombol delete
        LOADING.SetParm Me, 100
        
        Dim Aa As String
        Aa = MsgBox("Anda yakin ingin menghapus Data  '" & prjSysID.GetGridDefValue(dxDBGrid, "NO") & "' ?", vbYesNo, "Perhatian")
        
        If Aa = vbYes Then
            PrjSysTrig.DataDel conMain, "NO", "NO", "'" & prjSysID.GetGridDefValue(dxDBGrid, "NO") & "'"
            MenuGrid
            GridFill
        Else
            Exit Sub
        End If
        Exit Sub
    Case 1
        '=========
        '---Add---
        '=========
        If cmdList(1).Caption = "Add" Then
           With dxDBGrid.Dataset
               .Insert
                For i = 0 To 3
                    dxDBGrid.Columns(i).DisableEditor = False
                    dxDBGrid.Columns(i).Color = &HFFFFFF
                    cmdList(1).Caption = "Accept"
                Next i
            End With
        Else
            cmdList(1).Caption = "Add"
            dxDBGrid.Dataset.Refresh
        End If
    Case 2
        '==========
        '---Edit---
        '==========
        If cmdList(2).Caption = "Edit" Then
            For i = 5 To dxDBGrid.Columns.Count - 1
                dxDBGrid.Columns(i).DisableEditor = False
                dxDBGrid.Columns(i).Color = &HFFFFFF
                cmdList(2).Caption = "Accept"
            Next i
        Else
            For i = 0 To dxDBGrid.Columns.Count - 1
               dxDBGrid.Columns(i).DisableEditor = True
               dxDBGrid.Columns(i).Color = &HCBFAFE
               cmdList(2).Caption = "Edit"
            Next i
        End If
End Select
LOADING.SetParm Me, 100
Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
    LOADING.SetParm Me, 100
End Sub

Private Sub dxDBGrid_OnCustomDrawCell(ByVal hDC As Long, ByVal left As Single, ByVal top As Single, ByVal right As Single, ByVal bottom As Single, ByVal Node As DXDBGRIDLibCtl.IdxGridNode, ByVal Column As DXDBGRIDLibCtl.IdxGridColumn, ByVal Selected As Boolean, ByVal Focused As Boolean, ByVal NewItemRow As Boolean, Text As String, Color As Long, ByVal Font As stdole.IFontDisp, FontColor As Long, Alignment As DXDBGRIDLibCtl.ExAlignment, Done As Boolean)
Dim s As Integer
Dim stsRo As Boolean

On Error Resume Next

s = Node.RecNo 'Node.Values(dxDBGridROitem.Columns.ColumnByFieldName("NO").Index)
        If s Mod 2 = 0 Then
           Color = RGB(255, 255, 255)
        Else
           Color = &HE0E0E0                'RGB(255, 255, 222)
        End If
End Sub

Private Sub Form_Load()
lR = SetTopMostWindow(Me.hwnd, True)

On Error GoTo ErrorLabel
LOADING.Show
LOADING.SetParm Me, 25

'InitExample

'SrcBy = "True"
'fill combo reff
'cmbReff.Clear
'cmbReff.AddItem "Cari Nama"
'cmbReff.ListIndex = 0
LOADING.SetParm Me, 45

'fill image properti
'cmdList(0).Picture = LoadPicture(ImagePath("BTN_EDIT")) 'tombol Pilih
'cmdList(1).Picture = LoadPicture(ImagePath("BTN_ADD")) 'tombol TAMBAH
'cmdList(2).Picture = LoadPicture(ImagePath("BTN_WRITE")) 'tombol EDIT
'TabBack.PictureAlignment = xtpPictureTile
'TabBack.Picture = LoadPicture(ImagePath("PATERN1"))
Main
LOADING.SetParm Me, 45
prjSysID.CABANG cmbFilter(0)
prjSysID.Dept cmbFilter(1)
prjSysID.TTGROUP cmbFilter(2)

GridFill



LOADING.SetParm Me, 100

Exit Sub
ErrorLabel:
    If Err.Number <> 0 Then MsgBox CekError(Err.Number), vbCritical, "LG Error"
    LOADING.SetParm Me, 100

End Sub

Private Sub MenuGrid()
On Error Resume Next
    CLM1 = Array(Array("Clm0", "EmployeID", "KAR_ID", gedTextEdit, 0, 0, 120, 1, 1, 0, 0), _
                Array("Clm1", "Employe Name", "KAR_NM", gedTextEdit, 0, 0, 150, 1, 1, 0, 0), _
                Array("Clm2", "Cabang", "CAB_ID", gedLookupEdit, 0, 0, 100, 1, 1, 0, 0), _
                Array("Clm3", "Departmen ", "DEP_ID", gedLookupEdit, 0, 0, 120, 1, 1, 0, 0), _
                Array("Clm4", "Group", "GRP_ID", gedLookupEdit, 0, 0, 80, 1, 0, 0, 0), _
                Array("Clm5", "UPAH HARIAN", "PAY_DAY", gedCurrencyEdit, 0, 0, 100, 1, 2, 0, 6), _
                Array("Clm6", "SALARY", "PAY_MONTH", gedCurrencyEdit, 0, 0, 100, 1, 2, 0, 6), _
                Array("Clm7", "TUNJANGAN", "PAY_TUNJANGAN", gedCurrencyEdit, 0, 0, 100, 1, 2, 0, 6), _
                Array("Clm8", "TRANSPORT", "PAY_TRANPORT", gedCurrencyEdit, 0, 0, 100, 1, 2, 0, 6), _
                Array("Clm9", "MAKAN", "PAY_EAT", gedCurrencyEdit, 0, 0, 100, 1, 2, 0, 6), _
                Array("Clm10", "ENTERTAIN", "PAY_ENTERTAIN", gedCurrencyEdit, 0, 0, 100, 1, 3, 0, 6), _
                Array("Clm11", "BONUS", "BONUS", gedCurrencyEdit, 0, 0, 100, 1, 3, 0, 6))
            '--------0---------1--------2-----------3--------4--5--6---7--8--9
End Sub

Private Sub GridFill()
Dim strqry As String
On Error Resume Next
    MenuGrid
    strqry = "crud_salary(1,'" & cmbFilter(0).Text & "','" & cmbFilter(1).Text & "','" & cmbFilter(2).Text & "')"
    PrjSysGrid.GetGrid_Persensi dxDBGrid, CLM1, False, False, BND1, True, strqry, "KAR_ID", False
    LookupClm
    'dxDBGrid.Dataset.Open
End Sub
Private Sub LookupClm()
With dxDBGrid.Columns.ColumnByName("Clm2").LookupColumn
'== CABANG NAME ---
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT CAB_ID,CAB_NM FROM cabang " ' Like Join
            '.LookupDataset.Open
            .LookupKeyField = "CAB_ID"
            .LookupResultField = "CAB_NM"
            .ListColumns = "NAMA CABANG"
            .ListFieldName = "CAB_NM"
            .ListWidth = 200
            .DisplaySize = 400
End With

With dxDBGrid.Columns.ColumnByName("Clm3").LookupColumn
'== DEPT NAME ---
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT DEP_ID,DEP_NM FROM departemen " ' Like Join
            '.LookupDataset.Open
            .LookupKeyField = "DEP_ID"
            .LookupResultField = "DEP_NM"
            .ListColumns = "NAMA DEPT"
            .ListFieldName = "DEP_NM"
            .ListWidth = 200
            .DisplaySize = 400
End With

With dxDBGrid.Columns.ColumnByName("Clm4").LookupColumn
'== GROUP ---
            .LookupDataset.EnableControls
            '.LookupDataset.Close
            '.Dataset.Refresh
            .LookupDatasetType = dtADODataset
            .LookupDataset.ADODataset.ConnectionString = StrCon
            .LookupDataset.ADODataset.CursorLocation = clUseClient
            .LookupDataset.ADODataset.LockType = ltReadOnly         ' Check Permmission
            .LookupDataset.ADODataset.CommandType = cmdText
            .LookupDataset.ADODataset.CommandText = "SELECT TT_GRP_ID ,TT_GRP_NM FROM timetable_grp " ' Like Join
            '.LookupDataset.Open
            .LookupKeyField = "TT_GRP_ID"
            .LookupResultField = "TT_GRP_NM"
            .ListColumns = "GROUP TIME TABLE"
            .ListFieldName = "TT_GRP_NM"
            .ListWidth = 100
            .DisplaySize = 400
End With
dxDBGrid.Dataset.Open
End Sub

Sub InitExample()
 dxDBGrid.Event = 1 'EGOnCustomDrawCell
 dxDBGrid.EventEnabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
lR = SetTopMostWindow(Me.hwnd, False)
End Sub

Private Sub txtNm_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    'MenuGrid
    GridFill
End If
End Sub


