VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton save2 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6480
      TabIndex        =   18
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton save1 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6480
      TabIndex        =   17
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11160
      TabIndex        =   16
      Top             =   7320
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11160
      TabIndex        =   15
      Top             =   5880
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EDIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11160
      TabIndex        =   14
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CARI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11160
      TabIndex        =   13
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11160
      TabIndex        =   12
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Muhammad Aminudin\Desktop\Project\database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "akun"
      Top             =   7800
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   3255
      Left            =   3360
      OleObjectBlob   =   "Form1.frx":0014
      TabIndex        =   11
      Top             =   4320
      Width           =   7455
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   6360
      TabIndex        =   8
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   3600
      TabIndex        =   7
      Top             =   3720
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   4560
      TabIndex        =   5
      Top             =   2640
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   4560
      TabIndex        =   2
      Top             =   2160
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   4560
      TabIndex        =   1
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Saldo Kredit"
      Height          =   255
      Left            =   6720
      TabIndex        =   10
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Saldo Debit"
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Kelompok"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Nama Akun"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "No. Akun"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "CHART OF ACCOUNT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'mengkosongkan text
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = "0"
    Text4.Text = "0"
    
    'mengunci textbox saldo
    Text3.Enabled = False
    Text4.Enabled = False
    
    'memunculkan tombol simpan u/ menambah
    save1.Visible = True
    
    'pengaturan text1
    Text1.MaxLength = 6
    Text1.SetFocus
End Sub

Private Sub Command2_Click()
    'menghilangkan tombol simpan
    save1.Visible = False
    save2.Visible = False
    'mengkosongkan textbox
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Combo1.Text = ""
    
    'melakukan pencarian
    respon = vbYes
    While respon = vbYes
        respon = InputBox("Silahkan masukkan nomor akun!", "Pencarian")
        SRC = "No_Akun='" + respon + "'"
        Data1.Recordset.FindFirst SRC
        If Data1.Recordset.NoMatch Then
            respon = MsgBox("Akun tidak ditemukan!", vbCritical, "Tidak ditemukan")
        Else
            Text1.Text = Data1.Recordset!No_Akun
            Text2.Text = Data1.Recordset!Nama_Akun
            Text3.Text = Data1.Recordset!Saldo_Debit
            Text4.Text = Data1.Recordset!Saldo_Kredit
            Combo1.Text = Data1.Recordset!Kelompok
            respon = MsgBox("Akun ditemukan", vbInformation, "Notice")
        End If
    Wend
    
    'mengunci saldo
    Text3.Enabled = False
    Text4.Enabled = False
End Sub

Private Sub Command3_Click()
    'variabel
    vartext1 = Data1.Recordset!Saldo_Debit
    vartext2 = Data1.Recordset!Saldo_Kredit
    'jumlah dari var
    varsum = Val(vartext1) + Val(vartext2)
    'menampilkan recordset yg dipilih
    Text1.Text = Data1.Recordset!No_Akun
    Text2.Text = Data1.Recordset!Nama_Akun
    Text3.Text = Data1.Recordset!Saldo_Debit
    Text4.Text = Data1.Recordset!Saldo_Kredit
    Combo1.Text = Data1.Recordset!Kelompok
    'pengecekan
    If varsum > 0 Then
        respon = MsgBox("Saldo debit dan kredit harus nol!", vbCritical, "Gagal Edit")
    Else
        save2.Visible = True
    End If
End Sub

Private Sub Command4_Click()
    'variabel
    vartext1 = Data1.Recordset!Saldo_Debit
    vartext2 = Data1.Recordset!Saldo_Kredit
    'jumlah dari var
    varsum = Val(vartext1) + Val(vartext2)
    
    'pengecekan
    If varsum > 0 Then
        respon = MsgBox("Saldo debit dan kredit harus nol!", vbCritical, "Gagal Hapus")
    Else
        Data1.Recordset.Delete
        Data1.Refresh
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Combo1.Text = ""
        respon = MsgBox("Berhasil dihapus", vbInformation, "Notice")
    End If
End Sub

Private Sub Command5_Click()
    End
End Sub

Private Sub DBGrid1_Click()
        Text1.Text = Data1.Recordset!No_Akun
        Text2.Text = Data1.Recordset!Nama_Akun
        Text3.Text = Data1.Recordset!Saldo_Debit
        Text4.Text = Data1.Recordset!Saldo_Kredit
        Combo1.Text = Data1.Recordset!Kelompok
        Text3.Enabled = False
        Text4.Enabled = False
End Sub

Private Sub Form_Activate()
    Combo1.AddItem ("Aktiva Lancar")
    Combo1.AddItem ("Aktiva Tetap")
    Combo1.AddItem ("Hutang Lancar")
    Combo1.AddItem ("Hutang Tetap")
    Combo1.AddItem ("Modal")
    Combo1.AddItem ("Biaya")
    Combo1.AddItem ("Pendapatan")
End Sub

Private Sub Form_Load()
    Form1.WindowState = 2
End Sub

Private Sub save1_Click()
    'parameter
    vartext = Strings.Right(Text1.Text, 4)
    'variabel cari
    Cari1 = "No_Akun='" + "AL" + vartext + "'"
    Cari2 = "No_Akun='" + "AT" + vartext + "'"
    Cari3 = "No_Akun='" + "HL" + vartext + "'"
    Cari4 = "No_Akun='" + "HT" + vartext + "'"
    Cari5 = "No_Akun='" + "MD" + vartext + "'"
    Cari6 = "No_Akun='" + "BA" + vartext + "'"
    Cari7 = "No_Akun='" + "PD" + vartext + "'"
    'seleksi data
    Data1.Recordset.FindFirst Cari1
    If Data1.Recordset.NoMatch Then
        Data1.Recordset.FindFirst Cari2
        If Data1.Recordset.NoMatch Then
            Data1.Recordset.FindFirst Cari3
            If Data1.Recordset.NoMatch Then
                Data1.Recordset.FindFirst Cari4
                If Data1.Recordset.NoMatch Then
                    Data1.Recordset.FindFirst Cari5
                    If Data1.Recordset.NoMatch Then
                        Data1.Recordset.FindFirst Cari6
                        If Data1.Recordset.NoMatch Then
                            Data1.Recordset.FindFirst Cari7
                            If Data1.Recordset.NoMatch Then
                                'menambahkan data
                                Data1.Recordset.AddNew
                                Data1.Recordset!No_Akun = Text1.Text
                                Data1.Recordset!Nama_Akun = Text2.Text
                                Data1.Recordset!Kelompok = Combo1.Text
                                Data1.Recordset!Saldo_Debit = Text3.Text
                                Data1.Recordset!Saldo_Kredit = Text4.Text
                                Data1.Recordset.Update
                                'mengkosongkan textbox
                                Text1.Text = ""
                                Text2.Text = ""
                                Text3.Text = ""
                                Text4.Text = ""
                                Combo1.Text = ""
                                respon = MsgBox("Berhasil menambahkan akun!", vbInformation, "Notice")
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        respon = MsgBox("Nomor akun sudah tersedia!", vbCritical, "Gagal")
    End If
End Sub

Private Sub save2_Click()
    'variabel
    vartext1 = Data1.Recordset!Saldo_Debit
    vartext2 = Data1.Recordset!Saldo_Kredit
    'jumlah dari var
    varsum = Val(vartext1) + Val(vartext2)
    
    'pengecekan
    If varsum > 0 Then
        respon = MsgBox("Saldo debit dan kredit harus nol!", vbCritical, "Gagal Edit")
    Else
        Data1.Recordset.Edit
        Data1.Recordset!No_Akun = Text1.Text
        Data1.Recordset!Nama_Akun = Text2.Text
        Data1.Recordset!Kelompok = Combo1.Text
        Data1.Recordset!Saldo_Debit = Text3.Text
        Data1.Recordset!Saldo_Kredit = Text4.Text
        Data1.Recordset.Update
        respon = MsgBox("Berhasil edit!", vbInformation, "Notice")
    End If
End Sub
