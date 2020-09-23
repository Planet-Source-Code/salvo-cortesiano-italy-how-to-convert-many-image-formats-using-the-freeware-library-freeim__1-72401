VERSION 5.00
Begin VB.UserControl ShowImage 
   Alignable       =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   DataBindingBehavior=   1  'vbSimpleBound
   DataSourceBehavior=   1  'vbDataSource
   EditAtDesignTime=   -1  'True
   PropertyPages   =   "imgshow.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "imgshow.ctx":0014
   Begin VB.PictureBox immagine 
      AutoRedraw      =   -1  'True
      Height          =   1575
      Left            =   360
      ScaleHeight     =   1515
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Image temp 
      Height          =   615
      Left            =   2040
      Top             =   2640
      Width           =   495
   End
End
Attribute VB_Name = "ShowImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Dichiarazioni di eventi:
Event DblClick() 'MappingInfo=immagine,immagine,-1,DblClick
Attribute DblClick.VB_Description = "Viene generato quando si preme e si rilascia due volte in rapida successione un pulsante del mouse su un oggetto."
Event Click() 'MappingInfo=immagine,immagine,-1,Click
Attribute Click.VB_Description = "Viene generato quando si preme e quindi si rilascia un pulsante del mouse su un oggetto."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=immagine,immagine,-1,MouseMove
Attribute MouseMove.VB_Description = "Viene generato quando si sposta il mouse."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=immagine,immagine,-1,MouseDown
Attribute MouseDown.VB_Description = "Viene generato quando si preme il pulsante del mouse mentre lo stato attivo si trova su un oggetto."

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=immagine,immagine,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Restituisce o imposta lo stile del bordo di un oggetto."
   BorderStyle = immagine.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
   immagine.BorderStyle() = New_BorderStyle
   PropertyChanged "BorderStyle"
End Property

Private Sub immagine_Click()
   RaiseEvent Click
End Sub

Private Sub immagine_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub immagine_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub


'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=immagine,immagine,-1,Cls
Public Sub Cls()
Attribute Cls.VB_Description = "Cancella le immagini e il testo generati in fase di esecuzione da un form o da un controllo Image o PictureBox."
   immagine.Cls
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Set temp.Picture = PropBag.ReadProperty("Picture", Nothing)
   immagine.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
   immagine.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
End Sub

Private Sub UserControl_Resize()
immagine.left = 0
immagine.top = 0
immagine.Height = UserControl.Height
immagine.Width = UserControl.Width
Call showImageInfo
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   Call PropBag.WriteProperty("Picture", temp.Picture, Nothing)
   Call PropBag.WriteProperty("BorderStyle", immagine.BorderStyle, 1)
   Call PropBag.WriteProperty("BackColor", immagine.BackColor, &H8000000F)
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=immagine,immagine,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Restituisce o imposta il colore di sfondo utilizzato per la visualizzazione di testo e grafica in un oggetto."
   BackColor = immagine.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   immagine.BackColor = New_BackColor
   PropertyChanged "BackColor"
Call showImageInfo
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=temp,temp,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Restituisce o imposta un elemento grafico da visualizzare in un controllo."
   Set Picture = temp.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
   Set temp.Picture = New_Picture
   PropertyChanged "Picture"
Call showImageInfo
End Property

Public Sub showImageInfo()
   If temp.Picture <> UserControl.Picture Then
Dim irapp, irapp2
Dim iw As Integer, ih As Integer, ix As Integer, iy As Integer
   immagine.Cls
   irapp = temp.Width / temp.Height
   irapp2 = immagine.Width / immagine.Height
   If irapp >= 1 Then
      If irapp2 >= irapp Then
         ih = immagine.Height
         iw = immagine.Height * irapp
         iy = 0
         ix = (immagine.Width - iw) / 2
      Else
         iw = immagine.Width
         ih = immagine.Width / irapp
         ix = 0
         iy = (immagine.Height - ih) / 2
         End If
   Else
      If irapp2 <= irapp Then
         iw = immagine.Width
         ih = immagine.Width / irapp
         iy = (immagine.Height - ih) / 2
      Else
         ih = immagine.Height
         iw = immagine.Height * irapp
         iy = 0
         ix = (immagine.Width - iw) / 2
      End If
   End If
   
   immagine.PaintPicture temp.Picture, ix, iy, iw, ih
   End If

End Sub
Private Sub immagine_DblClick()
   RaiseEvent DblClick
End Sub

Public Function loadimg(nome As String) As Picture
temp.Picture = LoadPicture(nome)
Call showImageInfo
End Function


Public Property Get RegistrationCode() As String
   RegistrationCode = immagine.Tag
End Property

Public Property Let RegistrationCode(ByVal New_RegistrationCode As String)
   immagine.Tag = New_RegistrationCode
   PropertyChanged "RegistrationCode"
End Property
