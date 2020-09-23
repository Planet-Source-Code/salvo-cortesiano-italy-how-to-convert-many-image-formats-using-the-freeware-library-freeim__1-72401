VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Converted Free v1.0.2"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":151A
   ScaleHeight     =   5235
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbFormats 
      Height          =   330
      ItemData        =   "frmMain.frx":305A
      Left            =   1590
      List            =   "frmMain.frx":308D
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4140
      Width           =   1590
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   330
      Left            =   2385
      TabIndex        =   10
      Top             =   4680
      Width           =   1230
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   330
      Left            =   5070
      TabIndex        =   9
      Top             =   4680
      Width           =   1230
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   330
      Left            =   7845
      TabIndex        =   8
      Top             =   4680
      Width           =   1230
   End
   Begin VB.Frame frms 
      BackColor       =   &H80000005&
      Caption         =   "Browser Images"
      Height          =   3405
      Index           =   0
      Left            =   1575
      TabIndex        =   1
      Top             =   690
      Width           =   8460
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   3165
         Left            =   30
         ScaleHeight     =   3165
         ScaleWidth      =   8385
         TabIndex        =   2
         Top             =   195
         Width           =   8385
         Begin ImageConvertedFree.ShowImage prewImage 
            Height          =   2775
            Left            =   4590
            TabIndex        =   6
            Top             =   75
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   4895
            BorderStyle     =   0
            BackColor       =   -2147483643
         End
         Begin VB.FileListBox dirFiles 
            Height          =   1140
            Left            =   60
            Pattern         =   "*.bmp;*.jpg;*.tif;*.png;*.dib;*.gif;*.ico;*.pcx"
            System          =   -1  'True
            TabIndex        =   5
            Top             =   1965
            Width           =   4365
         End
         Begin VB.DirListBox dirDrive 
            Height          =   1530
            Left            =   45
            TabIndex        =   4
            Top             =   405
            Width           =   4395
         End
         Begin VB.DriveListBox drv 
            Height          =   330
            Left            =   45
            TabIndex        =   3
            Top             =   75
            Width           =   4410
         End
         Begin VB.Label lblFileName 
            BackColor       =   &H80000005&
            Caption         =   "##"
            Height          =   255
            Left            =   4620
            TabIndex        =   7
            Top             =   2880
            Width           =   3675
         End
      End
   End
   Begin VB.Label lbls 
      BackColor       =   &H80000005&
      Caption         =   "Select the formats"
      ForeColor       =   &H8000000D&
      Height          =   225
      Index           =   1
      Left            =   3315
      TabIndex        =   12
      Top             =   4200
      Width           =   2460
   End
   Begin VB.Label lbls 
      BackColor       =   &H80000005&
      Caption         =   "This is a freeware tool that interfaces with the DLL freeware 'FreeImage.dll' of SourceForge, to convert all konow Images format!"
      Height          =   525
      Index           =   0
      Left            =   2655
      TabIndex        =   0
      Top             =   240
      Width           =   7515
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2010
      Picture         =   "frmMain.frx":3118
      Top             =   150
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fFileName As String

Private Sub cmdAbout_Click()
    MsgBox "U can download the complete Project at: http://freeimage.sourceforge.net/", vbInformation, App.Title
End Sub

Private Sub cmdConvert_Click()
    Dim dib As Long: Dim bOK As Long
    Dim exts As String: Dim imgFormat As FREE_IMAGE_FORMAT
    Dim sFileName As String
    
    On Local Error GoTo ErrorConversion
    
    '/// Retrive the formats
    '/// This is only a DEMO, to implement hoter formats see the project
    exts = GetFilePath(fFileName, Only_Extension)
    
    sFileName = GetFilePath(fFileName, Only_FileName_no_Extension)
    
    Select Case exts
        Case "bmp", "dib": imgFormat = FIF_BMP
        Case "jpg": imgFormat = FIF_JPEG
        Case "gif": imgFormat = FIF_GIF
        Case "png": imgFormat = FIF_PNG
        Case "ico": imgFormat = FIF_ICO
        Case "tif": imgFormat = FIF_TIFF
        Case "iff": imgFormat = FIF_IFF
        Case "pcx": imgFormat = FIF_PCX
    End Select
    
    '/// Load a image
    dib = FreeImage_Load(imgFormat, fFileName, 0)
    
    Dim fType As String
    
    Select Case cmbFormats.List(cmbFormats.ListIndex)
        Case "FIF_BMP": fType = ".bmp"
        Case "FIF_JPEG": fType = ".jpg"
        Case "FIF_GIF": fType = ".gif"
        Case "FIF_PNG": fType = ".png"
        Case "FIF_ICO": fType = ".ico"
        Case "FIF_TIFF": fType = ".tif"
        Case "FIF_IFF": fType = ".iff"
        Case "FIF_PCX": fType = ".pcx"
    End Select
    
    If fType = "." & exts Then
            MsgBox "Sorry, change the convertion file Image!!", vbExclamation, App.Title
        Exit Sub
    End If
    
    '/// Save this image as PNG
    '/// parameters File type to be converted, file to be converted, new image name, image save options
    bOK = FreeImage_Save(cmbFormats.ItemData(cmbFormats.ListIndex), dib, App.Path + "\" + sFileName & fType, 0)
  
    '/// Unload the dib
    FreeImage_Unload (dib)
Exit Sub
ErrorConversion:
        MsgBox "Error #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub dirDrive_Change()
    On Error Resume Next
    dirFiles.Path = dirDrive.Path
    If dirFiles.ListCount > 0 Then
        dirFiles.Selected(0) = True
        lblFileName.Caption = dirFiles.Filename
        fFileName = dirDrive.List(dirDrive.ListIndex) + "\" + dirFiles.Filename
        prewImage.loadimg fFileName
    End If
End Sub

Private Sub dirFiles_DblClick()
    On Local Error GoTo ErrorHandler
    lblFileName.Caption = dirFiles.Filename
    fFileName = dirDrive.List(dirDrive.ListIndex) + "\" + dirFiles.Filename
    If FileExists(fFileName) = False Then
            MsgBox "Sorry, " & dirFiles.Filename & ", not found in the selected path!", vbExclamation, App.Title
        Exit Sub
    End If
    prewImage.loadimg fFileName
Exit Sub
ErrorHandler:
        MsgBox "Error: #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub


Private Sub dirFiles_KeyUp(KeyCode As Integer, Shift As Integer)
On Local Error GoTo ErrorHandler
    lblFileName.Caption = dirFiles.Filename
    fFileName = dirDrive.List(dirDrive.ListIndex) + "\" + dirFiles.Filename
    If FileExists(fFileName) = False Then
            MsgBox "Sorry, " & dirFiles.Filename & ", not found in the selected path!", vbExclamation, App.Title
        Exit Sub
    End If
    prewImage.loadimg fFileName
Exit Sub
ErrorHandler:
        MsgBox "Error: #" & Err.Number & ". " & Err.Description, vbExclamation, App.Title
    Err.Clear
End Sub


Private Sub drv_Change()
    On Error Resume Next
    dirDrive.Path = drv.Drive
End Sub

Private Sub Form_Initialize()
    '/// Init Controls XP/Vista Manifest
    '/// *****************************************************************
    Call InitCommonControlsVB
End Sub

Private Sub Form_Load()
    '/// Verify if the DLL is in the current Path
    '/// *****************************************************************
    If FileExists(App.Path + "\FreeImage.dll") = False Then
            MsgBox "Sorry, the FreeImage.dll not found in the current path!" & vbCr _
            & "Put the {FreeImage.dll} into current path before runnig this Application!", vbExclamation, App.Title
        Unload Me
    End If
    
    cmbFormats.ListIndex = 10
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmMain = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub


