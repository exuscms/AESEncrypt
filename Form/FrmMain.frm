VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Simple Zip"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9945
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6810
   ScaleWidth      =   9945
   StartUpPosition =   2  '화면 가운데
   Begin MSComctlLib.ListView Lstfile 
      Height          =   1215
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2143
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움체"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "이름"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "용량"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "확장명"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB1 
      Align           =   1  '위 맞춤
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   635
      ButtonWidth     =   2408
      ButtonHeight    =   582
      ToolTips        =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "새 파일"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "열기"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "풀기"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "압축 암호화"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "압축 복호화"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "암호화"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "복호화"
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Plaintext is hex"
      Height          =   255
      Left            =   -120
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   3360
      TabIndex        =   1
      Text            =   "Password Passphrase"
      Top             =   840
      Width           =   8055
   End
   Begin VB.ComboBox cboBlockSize 
      Appearance      =   0  '평면
      Height          =   300
      Left            =   1680
      Style           =   2  '드롭다운 목록
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.ComboBox cboKeySize 
      Appearance      =   0  '평면
      Height          =   300
      Left            =   0
      Style           =   2  '드롭다운 목록
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CDFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CDZip 
      Left            =   -360
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CDUnzip 
      Left            =   -360
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CDUnlock 
      Left            =   -360
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CDUnlock2 
      Left            =   -360
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "파일(&F)"
      Begin VB.Menu mnuNew 
         Caption         =   "새 압축파일(&N)"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "압축파일 열기(&O)"
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveAlz 
         Caption         =   "압축파일 저장(&S)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "종료(&Q)"
      End
   End
   Begin VB.Menu mnuList 
      Caption         =   "목록(&L)"
      Enabled         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "파일 추과(&A)"
      End
   End
   Begin VB.Menu mnuZip 
      Caption         =   "압축(Z)"
      Enabled         =   0   'False
      Begin VB.Menu mnuUnzip 
         Caption         =   "압축풀기(&U)"
      End
   End
   Begin VB.Menu mnuAES 
      Caption         =   "암호화(&A)"
      Begin VB.Menu mnuSave 
         Caption         =   "압축파일 암호화(&S)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuUnlock 
         Caption         =   "압축파일 복호화(&U)"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnufileAES 
         Caption         =   "파일 암호화(&A)"
      End
      Begin VB.Menu mnuOnlySave 
         Caption         =   "파일 복호화(&S)"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const bits = 1024
Dim fname As String
Dim info As file_info
Dim f1, f2, f3 As Integer
Dim buffer2() As Byte
Dim Buffer(1 To bits) As Byte
#Const SUPPORT_LEVEL = 0
Private WithEvents TaskBarList As ITaskBarList3
Attribute TaskBarList.VB_VarHelpID = -1
Private m_Rijndael As New cRijndael
Private Sub Form_Load()
    Set TaskBarList = New ITaskBarList3
    cboBlockSize.AddItem "AES-128"
    cboBlockSize.ItemData(cboBlockSize.NewIndex) = 128
#If SUPPORT_LEVEL = 0 Then
    cboBlockSize.Enabled = False
#Else
#If SUPPORT_LEVEL = 2 Then
    cboBlockSize.AddItem "160 Bit"
    cboBlockSize.ItemData(cboBlockSize.NewIndex) = 160
    cmdSizeTest.Visible = True
#End If
    cboBlockSize.AddItem "AES-192"
    cboBlockSize.ItemData(cboBlockSize.NewIndex) = 192
#If SUPPORT_LEVEL = 2 Then
    cboBlockSize.AddItem "AES-224"
    cboBlockSize.ItemData(cboBlockSize.NewIndex) = 224
#End If
    cboBlockSize.AddItem "AES-256"
    cboBlockSize.ItemData(cboBlockSize.NewIndex) = 256
#End If
    cboKeySize.AddItem "AES-128"
    cboKeySize.ItemData(cboKeySize.NewIndex) = 128
#If SUPPORT_LEVEL = 2 Then
    cboKeySize.AddItem "AES-160"
    cboKeySize.ItemData(cboKeySize.NewIndex) = 160
#End If
    cboKeySize.AddItem "AES-192"
    cboKeySize.ItemData(cboKeySize.NewIndex) = 192
#If SUPPORT_LEVEL = 2 Then
    cboKeySize.AddItem "AES-224"
    cboKeySize.ItemData(cboKeySize.NewIndex) = 224
#End If
    cboKeySize.AddItem "AES-256"
    cboKeySize.ItemData(cboKeySize.NewIndex) = 256
    cboBlockSize.ListIndex = 0
    cboKeySize.ListIndex = 0
    txtPassword = ""
    Status = ""

End Sub

Private Sub Form_Resize()
On Error Resume Next
cboKeySize.Top = TB1.Height
cboBlockSize.Top = TB1.Height
txtPassword.Top = TB1.Height
txtPassword.Width = Me.ScaleWidth - (cboKeySize.Width + cboBlockSize.Width)
Lstfile.Top = TB1.Height + txtPassword.Height + 30
Lstfile.Width = Me.ScaleWidth
Lstfile.Height = Me.ScaleHeight - (TB1.Height + txtPassword.Height + 30)
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuAdd_Click()
On Error Resume Next
CDZip.ShowOpen
Dim c, w

If Not CDZip.Filename = "" Then
    c = CDZip.Filename
    
    Do While Not InStr(c, "\") = "0"
    DoEvents
        If InStr(1, c, "\") Then
            w = InStr(1, c, "\")
            If w >= 2 Then
                c = Right(CDZip.Filename, Len(c) - (w - 1))
            ElseIf w = 1 Then
                c = Right(CDZip.Filename, Len(c) - (w))
            End If
        End If
    Loop
    
    Lstfile.ListItems.Add Lstfile.ListItems.Count + 1, CDZip.Filename, c
    Lstfile.ListItems.Item(Lstfile.ListItems.Count).SubItems(1) = Format((FileLen(CDZip.Filename)), "#,###") & "Kb"
End If
End Sub

Public Sub mnufileAES_Click()
'On Error Resume Next
Dim pass()    As Byte
Dim KeyBits   As Long
Dim BlockBits As Long
Dim k, j, rest As Integer

If txtPassword = "" Then

    MsgBox "비밀번호를 입력하시오", vbCritical, "오류"
    txtPassword.SetFocus
    Exit Sub

End If

CDUnlock.DialogTitle = "암호화할 단일파일"
CDUnlock.ShowOpen

CDUnlock2.DialogTitle = "저장할 경로 및 이름"
CDUnlock2.ShowOpen

If Not CDUnlock.Filename = "" And Not CDUnlock2.Filename = "" Then
    FrmZip.Show
    FrmMain.Hide
    TaskBarList.SetProgressState FrmZip.hwnd, 2 ^ (0)
                KeyBits = cboKeySize.ItemData(cboKeySize.ListIndex)
                BlockBits = cboBlockSize.ItemData(cboBlockSize.ListIndex)
                pass = GetPassword

                Status = "Encrypting File"
#If SUPPORT_LEVEL Then
                m_Rijndael.SetCipherKey pass, KeyBits, BlockBits
                'FrmZip.Caption = "암호화중...(BlockBits)" & FileName1
                m_Rijndael.FileEncrypt Fileset(CDUnlock.Filename), Fileset(CDUnlock2.Filename), BlockBits
#Else
                m_Rijndael.SetCipherKey pass, KeyBits
                'FrmZip.Caption = "암호화중..." & FileName1
                m_Rijndael.FileEncrypt Fileset(CDUnlock.Filename), Fileset(CDUnlock2.Filename)
#End If
                Status = ""
    TaskBarList.SetProgressState FrmZip.hwnd, 2 ^ (-1)
    Unload FrmZip
    FrmMain.Show

End If
End Sub

Public Sub mnuNew_Click()
mnuZip.Enabled = False
mnuList.Enabled = True
mnuSave.Enabled = True
mnuSaveAlz.Enabled = True
Me.Caption = "Secret Zip - " & "새파일.zip"
Lstfile.ListItems.Clear
End Sub

Private Sub mnuOnlySave_Click()
Dim pass()    As Byte
Dim KeyBits   As Long
Dim BlockBits As Long

If txtPassword = "" Then

    MsgBox "비밀번호를 입력하시오", vbCritical, "오류"
    txtPassword.SetFocus
    Exit Sub

End If

CDUnlock.DialogTitle = "암호화를 해제할 파일"
CDUnlock.ShowOpen

If CDUnlock.Filename = "" Then
    Exit Sub
End If

CDUnlock2.DialogTitle = "해제된 파일을 풀 경로"
CDUnlock2.ShowOpen

If Not CDUnlock.Filename = "" And Not CDUnlock2.Filename = "" Then

    FrmZip.Show
    'FrmMain.Hide
    TaskBarList.SetProgressState FrmZip.hwnd, 2 ^ (0)
    KeyBits = cboKeySize.ItemData(cboKeySize.ListIndex)
    BlockBits = cboBlockSize.ItemData(cboBlockSize.ListIndex)
    pass = GetPassword

    Status = "Decrypting File"

    #If SUPPORT_LEVEL Then
                m_Rijndael.SetCipherKey pass, KeyBits, BlockBits
                m_Rijndael.FileDecrypt Fileset(CDUnlock2.Filename), Fileset(CDUnlock.Filename), BlockBits
                FrmZip.Caption = "복호화중...(BlockBits)" & (CDUnlock2.Filename)
    #Else
                m_Rijndael.SetCipherKey pass, KeyBits
                m_Rijndael.FileDecrypt Fileset(CDUnlock2.Filename), Fileset(CDUnlock.Filename)
                FrmZip.Caption = "복호화중..." & CDUnlock2.Filename
    #End If
                Status = ""
    TaskBarList.SetProgressState FrmZip.hwnd, 2 ^ (-1)
    FrmZip.Hide
    FrmMain.Show
End If
End Sub

Public Sub mnuOpen_Click()
On Error GoTo Decrypt

Dim file As String
Dim f1, f2, k As Integer

CDFile.ShowOpen

If Not CDFile.Filename = "" Then

    fname = CDFile.Filename
    file = CDFile.Filename & ".name"
    If FileLen(file) = "0" Then Exit Sub
    f2 = FreeFile
    Open file For Random As f2
    Lstfile.ListItems.Clear
    Do
    DoEvents
    k = k + 1
    Get f2, k, info
    If Not info.File_Name = "" Then
        Lstfile.ListItems.Add Lstfile.ListItems.Count + 1, , (info.File_Name)
    End If
    Loop Until EOF(f2)
    mnuList.Enabled = False
    mnuZip.Enabled = True
    
End If
Exit Sub
Decrypt:
MsgBox "복호화가 필요한 파일이거나 잘못된 파일입니다.", vbCritical, "오류"
End Sub

Private Sub mnuQuit_Click()
End
End Sub

Public Sub mnuSave_Click()
On Error Resume Next
Dim pass()    As Byte
Dim KeyBits   As Long
Dim BlockBits As Long
Dim k, j, rest As Integer
Dim size, count_size, current_pos As Long
Dim FileName1, FileName2, FileName3 As String

If txtPassword = "" Then

    MsgBox "비밀번호를 입력하시오", vbCritical, "오류"
    txtPassword.SetFocus
    Exit Sub

End If

CDFile.ShowSave

If Not CDFile.Filename = "" Then
    
    FrmZip.Show
    FrmMain.Hide

    current_pos = 1

    FileName1 = CDFile.Filename & ".aes" & cboKeySize.ItemData(cboKeySize.ListIndex)
    FileName2 = CDFile.Filename & ".aes" & cboKeySize.ItemData(cboKeySize.ListIndex) & ".name"

    f1 = FreeFile
    Open FileName1 For Binary As f1
    f2 = FreeFile
    Open FileName2 For Random As f2

    For k = 1 To Lstfile.ListItems.Count
    DoEvents

    FileName3 = Lstfile.ListItems(k).Key
    
    c = Lstfile.ListItems(k).Key
    
    Do While Not InStr(c, "\") = "0"
    DoEvents
        If InStr(1, c, "\") Then
            w = InStr(1, c, "\")
            If w >= 2 Then
                c = Right(c, Len(c) - (w - 1))
            ElseIf w = 1 Then
                c = Right(c, Len(c) - (w))
            End If
        End If
    Loop
    
    FrmZip.LabName.Caption = c
    FrmZip.Caption = Lstfile.ListItems(k).Key

    f3 = FreeFile
    Open FileName3 For Binary As f3

    FrmZip.PBFile.Value = Replace(Format(Str(k) / Str(Lstfile.ListItems.Count), "0%"), "%", "")
    FrmZip.LabPer.Caption = Format(Str(k) / Str(Lstfile.ListItems.Count), "0%")

    size = FileLen(FileName3)

    info.File_Name = Lstfile.ListItems(k).Text
    info.file_size = size
    count_size = size
    info.file_pos = current_pos

    While size >= bits
    DoEvents
    FrmZip.LabBuffer.Caption = Val(Replace(Format((count_size - size) / count_size, "0%"), "%", "")) & "%"
    FrmZip.PBFile2.Value = Val(Replace(Format((count_size - size) / count_size, "0%"), "%", ""))
    TaskBarList.SetProgressValue FrmZip.hwnd, Val(Replace(Format((count_size - size) / count_size, "0%"), "%", "")), 100
    size = size - bits
    Get f3, , Buffer
    Put f1, , Buffer
    Wend

    If size > 0 Then
        FrmZip.LabBuffer.Caption = size
        ReDim buffer2(1 To size)
        Get f3, , buffer2
        Put f1, , buffer2
    End If
    
    Put 2, k + 1, info
    Close f3
    current_pos = current_pos + info.file_size
    
    Next
    
    Close (f1)
    Close (f2)
                TaskBarList.SetProgressState FrmZip.hwnd, 2 ^ (0)
                KeyBits = cboKeySize.ItemData(cboKeySize.ListIndex)
                BlockBits = cboBlockSize.ItemData(cboBlockSize.ListIndex)
                pass = GetPassword
                Status = "Encrypting File"
#If SUPPORT_LEVEL Then
                m_Rijndael.SetCipherKey pass, KeyBits, BlockBits
                FrmZip.Caption = "암호화중...(BlockBits)" & FileName1
                m_Rijndael.FileEncrypt Fileset(FileName1), Fileset(FileName1), BlockBits
                FrmZip.Caption = "암호화중...(BlockBits)" & FileName2
                m_Rijndael.FileEncrypt Fileset(FileName2), Fileset(FileName2), BlockBits
#Else
                m_Rijndael.SetCipherKey pass, KeyBits
                FrmZip.Caption = "암호화중..." & FileName1
                m_Rijndael.FileEncrypt Fileset(FileName1), Fileset(FileName1)
                FrmZip.Caption = "암호화중..." & FileName2
                m_Rijndael.FileEncrypt Fileset(FileName2), Fileset(FileName2)
#End If
                Status = ""
    
    TaskBarList.SetProgressState FrmZip.hwnd, 2 ^ (-1)
    Unload FrmZip
    FrmMain.Show

End If
End Sub

Private Function HexDisplayRev(TheString As String, data() As Byte) As Long
    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim d As Long
    Dim n As Long
    Dim Data2() As Byte

    n = 2 * Len(TheString)
    Data2 = TheString

    ReDim data(n \ 4 - 1)

    d = 0
    i = 0
    j = 0
    Do While j < n
        c = Data2(j)
        Select Case c
        Case 48 To 57    '"0" ... "9"
            If d = 0 Then   'high
                d = c
            Else            'low
                data(i) = (c - 48) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        Case 65 To 70   '"A" ... "F"
            If d = 0 Then   'high
                d = c - 7
            Else            'low
                data(i) = (c - 55) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        Case 97 To 102  '"a" ... "f"
            If d = 0 Then   'high
                d = c - 39
            Else            'low
                data(i) = (c - 87) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        End Select
        j = j + 2
    Loop
    n = i
    If n = 0 Then
        Erase data
    Else
        ReDim Preserve data(n - 1)
    End If
    HexDisplayRev = n

End Function

Private Function GetPassword() As Byte()
    Dim data() As Byte

    If Check1.Value = 0 Then
        data = StrConv(txtPassword.Text, vbFromUnicode)
        ReDim Preserve data(31)
    Else
        If HexDisplayRev(txtPassword.Text, data) <> (cboKeySize.ItemData(cboKeySize.ListIndex) \ 8) Then
            data = StrConv(txtPassword.Text, vbFromUnicode)
            ReDim Preserve data(31)
        End If
    End If
    GetPassword = data
End Function

Private Sub mnuSaveAlz_Click()
On Error Resume Next
Dim pass()    As Byte
Dim KeyBits   As Long
Dim BlockBits As Long
Dim k, j, rest As Integer
Dim size, count_size, current_pos As Long
Dim FileName1, FileName2, FileName3 As String

CDFile.ShowSave

If Not CDFile.Filename = "" Then

    FrmZip.Show
    FrmMain.Hide

    current_pos = 1

    FileName1 = CDFile.Filename
    FileName2 = CDFile.Filename & ".name"

    f1 = FreeFile
    Open FileName1 For Binary As f1
    f2 = FreeFile
    Open FileName2 For Random As f2

    For k = 1 To Lstfile.ListItems.Count
    DoEvents

    FileName3 = Lstfile.ListItems(k).Key
    FrmZip.LabName.Caption = Lstfile.ListItems(k).Key
    FrmZip.Caption = "압축중..." & Lstfile.ListItems(k).Key

    f3 = FreeFile
    Open FileName3 For Binary As f3

    FrmZip.PBFile.Value = Replace(Format(Str(k) / Str(Lstfile.ListItems.Count), "0%"), "%", "")
    FrmZip.LabPer.Caption = Format(Str(k) / Str(Lstfile.ListItems.Count), "0%")

    size = FileLen(FileName3)

    info.File_Name = Lstfile.ListItems(k).Text
    info.file_size = size
    count_size = size
    info.file_pos = current_pos

    While size >= bits
    DoEvents
    FrmZip.LabBuffer.Caption = Val(Replace(Format((count_size - size) / count_size, "0%"), "%", "")) & "%"
    FrmZip.PBFile2.Value = Val(Replace(Format((count_size - size) / count_size, "0%"), "%", ""))
    TaskBarList.SetProgressValue FrmZip.hwnd, Val(Replace(Format((count_size - size) / count_size, "0%"), "%", "")), 100
    size = size - bits
    Get f3, , Buffer
    Put f1, , Buffer
    Wend

    If size > 0 Then
        FrmZip.LabBuffer.Caption = size
        ReDim buffer2(1 To size)
        Get f3, , buffer2
        Put f1, , buffer2
    End If
    
    Put 2, k + 1, info
    Close f3
    current_pos = current_pos + info.file_size
    
    Next
    
    Close (f1)
    Close (f2)
    
    Unload FrmZip
    FrmMain.Show

End If
End Sub

Public Sub mnuUnlock_Click()
Dim pass()    As Byte
Dim KeyBits   As Long
Dim BlockBits As Long

If txtPassword = "" Then

    MsgBox "비밀번호를 입력하시오", vbCritical, "오류"
    txtPassword.SetFocus
    Exit Sub

End If

CDUnlock.DialogTitle = "암호화를 해제할 파일"
CDUnlock.ShowOpen

If CDUnlock.Filename = "" Then
    Exit Sub
End If

CDUnlock2.DialogTitle = "해제된 파일을 풀 경로"
CDUnlock2.ShowOpen

If Not CDUnlock.Filename = "" And Not CDUnlock2.Filename = "" Then

    FrmZip.Show
    'FrmMain.Hide
    TaskBarList.SetProgressState FrmZip.hwnd, 2 ^ (0)
    KeyBits = cboKeySize.ItemData(cboKeySize.ListIndex)
    BlockBits = cboBlockSize.ItemData(cboBlockSize.ListIndex)
    pass = GetPassword

    Status = "Decrypting File"

    #If SUPPORT_LEVEL Then
                m_Rijndael.SetCipherKey pass, KeyBits, BlockBits
                m_Rijndael.FileDecrypt Fileset(CDUnlock2.Filename), Fileset(CDUnlock.Filename), BlockBits
                FrmZip.Caption = "복호화중...(BlockBits)" & (CDUnlock2.Filename)
                m_Rijndael.FileDecrypt Fileset(CDUnlock2.Filename & ".name"), Fileset(CDUnlock.Filename & ".name"), BlockBits
                FrmZip.Caption = "복호화중...(BlockBits)" & CDUnlock2.Filename & ".name"
    #Else
                m_Rijndael.SetCipherKey pass, KeyBits
                m_Rijndael.FileDecrypt Fileset(CDUnlock2.Filename), Fileset(CDUnlock.Filename)
                FrmZip.Caption = "복호화중..." & CDUnlock2.Filename
                m_Rijndael.FileDecrypt Fileset(CDUnlock2.Filename & ".name"), Fileset(CDUnlock.Filename & ".name")
                FrmZip.Caption = "복호화중..." & CDUnlock2.Filename & ".name"
    #End If
                Status = ""
    TaskBarList.SetProgressState FrmZip.hwnd, 2 ^ (-1)
    FrmZip.Hide
    FrmMain.Show
End If
End Sub

Public Sub mnuUnzip_Click()
On Error Resume Next
Dim w, c, k As String
Dim current_file, f1, f2, f3 As Integer
Dim size, count_size, current_pos As Long
Dim File1, File2, File3 As String

CDUnzip.ShowSave
k = ""
If Not CDUnzip.Filename = "" Then

    FrmZip.Show
    FrmMain.Hide
    File1 = fname
    File2 = fname & ".name"

    FrmZip.LabName.Caption = File3
    
    current_pos = 1
    f1 = FreeFile
    Open File1 For Binary As f1
    f2 = FreeFile
    Open File2 For Random As f2

    current_file = 0
    While Not (EOF(f2))
    current_file = current_file + 1
    '''''''''''''''''''''''''''''''
    FrmZip.PBFile.Value = Replace(Format(Str(k) / Str(Lstfile.ListItems.Count), "0%"), "%", "")
    FrmZip.LabPer.Caption = Format(Str(k) / Str(Lstfile.ListItems.Count), "0%")
    '''''''''''''''''''''''''''''''
    Get f2, current_file, info

    c = CDUnzip.Filename
    If k = "" Then
    Do While Not InStr(c, "\") = "0"
    DoEvents
        If InStr(1, c, "\") Then
            w = InStr(1, c, "\")
            If w >= 2 Then
                c = Right(v, Len(c) - (w - 1))
            ElseIf w = 1 Then
                c = Right(v, Len(c) - (w))
            End If
        End If
    Loop
    
    k = Left(CDUnzip.Filename, Len(CDUnzip.Filename) - c)
    End If
    
    FrmZip.LabName = info.File_Name
    FrmZip.Caption = "압축푸는중..." & info.File_Name
    
    File3 = k & info.File_Name

    f3 = FreeFile

    Open File3 For Binary As f3

    size = info.file_size
    count_size = size
    
    While size > bits
    DoEvents
    size = size - bits
    FrmZip.LabBuffer.Caption = Val(Replace(Format((count_size - size) / count_size, "0%"), "%", "")) & "%"
    FrmZip.PBFile2.Value = Replace(Format((count_size - size) / count_size, "0%"), "%", "")
    TaskBarList.SetProgressValue FrmZip.hwnd, Val(Replace(Format((count_size - size) / count_size, "0%"), "%", "")), 100
    Get f1, , Buffer
    Put f3, , Buffer
    Wend

If size > 0 Then
    ReDim buffer2(1 To size) As Byte
    Get f1, , buffer2
    Put f3, , buffer2
End If

Close f3
Wend

Unload FrmZip
FrmMain.Show

End If
End Sub

Private Sub Tb1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1: mnuNew_Click
    Case 3: mnuOpen_Click
    Case 4: mnuUnzip_Click
    Case 6:
        If Lstfile.ListItems.Count = 0 Then
            MsgBox "압축할 파일이 없습니다", vbCritical, "오류"
        Else
            mnuSave_Click
        End If
    Case 7: mnuUnlock_Click
    Case 9: mnufileAES_Click
    Case 10: mnuOnlySave_Click
End Select
End Sub

Private Sub Lstfile_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
Dim i As Integer
Dim v
Dim w, c, a, b As String

If mnuList.Enabled = False Then
    mnuNew_Click
End If

If mnuList.Enabled = True Then

    For Each v In data.Files
    
    c = v
    
    Do While Not InStr(c, "\") = "0"
    DoEvents
        If InStr(1, c, "\") Then
            w = InStr(1, c, "\")
            If w >= 2 Then
                c = Right(v, Len(c) - (w - 1))
            ElseIf w = 1 Then
                c = Right(v, Len(c) - (w))
            End If
        End If
    Loop
    
    a = c
    
    Do While Not InStr(a, ".") = "0"
    DoEvents
        If InStr(1, a, ".") Then
            b = InStr(1, a, ".")
            If b >= 2 Then
                a = Right(v, Len(a) - (b - 1))
            ElseIf w = 1 Then
                a = Right(v, Len(a) - (b))
            End If
        End If
    Loop
    
                Lstfile.ListItems.Add Lstfile.ListItems.Count + 1, v, c
                Lstfile.ListItems.Item(Lstfile.ListItems.Count).SubItems(1) = Format((FileLen(v)), "#,###") & "Kb"
                Lstfile.ListItems.Item(Lstfile.ListItems.Count).SubItems(2) = "." & a

    Next
End If
End Sub

