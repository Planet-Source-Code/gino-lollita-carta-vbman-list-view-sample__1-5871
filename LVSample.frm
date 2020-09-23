VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Begin VB.Form LVSample 
   Caption         =   "ListView Sample"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   Icon            =   "LVSample.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdaptColumn2NoHeader 
      Caption         =   "Adatta colonna 2 senza tenere conto dell'intestazione"
      Height          =   540
      Left            =   5850
      TabIndex        =   26
      Top             =   4455
      Width           =   2460
   End
   Begin VB.CommandButton cmdAdaptColumn2 
      Caption         =   "Adatta colonna 2 tenendo conto dell'intestazione"
      Height          =   540
      Left            =   5850
      TabIndex        =   25
      Top             =   3885
      Width           =   2460
   End
   Begin VB.CommandButton cmdChangeRowBackColor 
      Caption         =   "Cambia colore sfondo righe"
      Height          =   450
      Left            =   5850
      TabIndex        =   24
      Top             =   5505
      Width           =   2460
   End
   Begin VB.CommandButton cmdChangeRowTextColor 
      Caption         =   "Cambia colore testo righe"
      Height          =   450
      Left            =   5850
      TabIndex        =   23
      Top             =   5025
      Width           =   2460
   End
   Begin VB.CommandButton cmdNoDataImage 
      Caption         =   "Immagine Nessun dato"
      Height          =   450
      Left            =   2640
      TabIndex        =   22
      Top             =   5505
      Width           =   2460
   End
   Begin VB.CommandButton cmdNoImage 
      Caption         =   "Nessuna immagine"
      Height          =   450
      Left            =   2640
      TabIndex        =   21
      Top             =   5025
      Width           =   2460
   End
   Begin VB.CommandButton cmdTiledImage 
      Caption         =   "Immagine affiancata"
      Height          =   450
      Left            =   2640
      TabIndex        =   20
      Top             =   4545
      Width           =   2460
   End
   Begin VB.CommandButton cmdCenteredImage 
      Caption         =   "Immagine centrata"
      Height          =   450
      Left            =   2640
      TabIndex        =   19
      Top             =   4065
      Width           =   2460
   End
   Begin VB.CommandButton cmdGetColumns 
      Caption         =   "Posizione originale colonne"
      Height          =   450
      Left            =   90
      TabIndex        =   18
      Top             =   5505
      Width           =   2460
   End
   Begin VB.CommandButton cmdSelectedRows 
      Caption         =   "Lista righe con check"
      Height          =   450
      Left            =   90
      TabIndex        =   17
      Top             =   5025
      Width           =   2460
   End
   Begin VB.CommandButton cmdIndentItem 
      Caption         =   "Indenta la riga corrente"
      Height          =   450
      Left            =   90
      TabIndex        =   16
      Top             =   4545
      Width           =   2460
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Togli icone ordinamento"
      Height          =   450
      Left            =   2640
      TabIndex        =   15
      Top             =   3480
      Width           =   2460
   End
   Begin VB.CommandButton cmdAddHeaderPics 
      Caption         =   "Aggiungi icone ordinamento"
      Height          =   450
      Left            =   105
      TabIndex        =   14
      Top             =   3480
      Width           =   2460
   End
   Begin VB.CheckBox chkImgPosition 
      Caption         =   "Bitmap ordinamento a destra"
      Height          =   330
      Left            =   5895
      TabIndex        =   13
      Top             =   3465
      Width           =   2430
   End
   Begin VB.CheckBox chkBitmapsOnAllColumns 
      Caption         =   "Bitmaps su tutte le colonne"
      Height          =   345
      Left            =   5895
      TabIndex        =   12
      Top             =   1935
      Width           =   2430
   End
   Begin VB.CheckBox chkStrikeOutFont 
      Caption         =   "Intestazioni in barrato"
      Height          =   345
      Left            =   5895
      TabIndex        =   11
      Top             =   3075
      Width           =   2430
   End
   Begin VB.CheckBox chkUnderlineFont 
      Caption         =   "Intestazioni in sottolineato"
      Height          =   345
      Left            =   5895
      TabIndex        =   10
      Top             =   2820
      Width           =   2430
   End
   Begin VB.CheckBox chkItalicFont 
      Caption         =   "Intestazioni in corsivo"
      Height          =   345
      Left            =   5895
      TabIndex        =   9
      Top             =   2565
      Width           =   2430
   End
   Begin VB.CheckBox chkBoldFont 
      Caption         =   "Intestazioni in grassetto"
      Height          =   345
      Left            =   5895
      TabIndex        =   8
      Top             =   2295
      Width           =   2430
   End
   Begin VB.CheckBox chkCheckBoxes 
      Caption         =   "Checkboxes"
      Height          =   345
      Left            =   5895
      TabIndex        =   7
      Top             =   1665
      Width           =   2430
   End
   Begin VB.CheckBox chkHeaderDragDrop 
      Caption         =   "Drag && drop intestazioni"
      Height          =   345
      Left            =   5895
      TabIndex        =   6
      Top             =   1410
      Width           =   2430
   End
   Begin VB.CheckBox chkHeaderHotTrack 
      Caption         =   "Tracking delle intestazioni"
      Height          =   345
      Left            =   5895
      TabIndex        =   5
      Top             =   1155
      Width           =   2430
   End
   Begin VB.CheckBox chkTrackSelect 
      Caption         =   "Tracking delle righe"
      Height          =   345
      Left            =   5895
      TabIndex        =   4
      Top             =   900
      Width           =   2430
   End
   Begin VB.CheckBox chkFullRowSelection 
      Caption         =   "Selezione riga intera"
      Height          =   345
      Left            =   5895
      TabIndex        =   3
      Top             =   630
      Width           =   2430
   End
   Begin VB.CheckBox chkGridLines 
      Caption         =   "Linee di griglia"
      Height          =   345
      Left            =   5895
      TabIndex        =   2
      Top             =   375
      Width           =   2430
   End
   Begin VB.CheckBox chkFlatHeaders 
      Caption         =   "Intestazioni piatte"
      Height          =   345
      Left            =   5895
      TabIndex        =   1
      Top             =   120
      Width           =   2430
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   3330
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   5874
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   5220
      Top             =   4170
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   12632256
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   22
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":035E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":06B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":09CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":0CE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":0FFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":1318
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":1632
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":194C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":1C66
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":1F80
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":229A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":25B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":28CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":2BE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":2F02
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":321C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":3536
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":3850
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":3B6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":3E84
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "LVSample.frx":419E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "LVSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Programma:    ListView Sample
' Autore:       Roberto Cappuccio
' Indirizzo:    roberto.cappuccio@nova.bz.it
' Data:         25/02/1999
' Note:
' Questo programma di esempio contiene tutto ciò che riesco
' a fare attualmente con il controllo ListView standard.
' Il codice proviene dalle fonti più disparate per cui mi
' è impossibile elencare qui tutti coloro che hanno contribuito
' a fare sì che la ListView non abbia più segreti... (o quasi)
' Dato che, in ogni caso, almeno il 60% del codice proviene
' da Brad Martinez, ritengo sia il caso di ringraziare almeno
' lui e tutti quelli che contribuiscono al sito VBNET
' che consiglio a tutti di visitare quanto prima:
'
'               http://www.mvps.org/vbnet/
'
' Nel caso qualcuno riesca a trovare altre capacità nascoste
' del controllo ListView (come ad esempio quella di colorare
' singole righe invece che tutto l'insieme degli items) il
' mio indirizzo è riportato sopra. Grazie.

Private imgPosition As Long
Public hHeaderFont As Long

Private Sub Form_Load()
    Dim i As Integer
    Dim itmX As ListItem
    Dim imgX As ListImage
    
    'Crea la listview e la riempie con dei dati
    With ListView1
        .View = lvwReport
        .SmallIcons = ImageList1.Object
        .ColumnHeaders.Add , "x1", "Colonna numero 1    "
        .ColumnHeaders.Item(1).Width = 1520
        .ColumnHeaders.Add , "x2", "Colonna numero 2    "
        .ColumnHeaders.Item(2).Width = 1520
        .ColumnHeaders.Add , "x3", "Colonna numero 3    "
        .ColumnHeaders.Item(3).Width = 1520
        For i = 1 To 19
            Set itmX = .ListItems.Add(, "key" & i, "main item" & i, , 7)
            itmX.SubItems(1) = "col1 subitem" & i
            itmX.SubItems(2) = "col2 subitem" & i
        Next i
    End With

    Call InitComctl32(ICC_LISTVIEW_CLASSES)

    'Mette a falso tutte le checkboxes delle opzioni
    'quindi inizialmente la lista ha un aspetto piuttosto normale...
    chkGridLines = 0
    chkFlatHeaders = 0
    chkFullRowSelection = 0
    chkTrackSelect = 0
    chkHeaderHotTrack = 0
    chkHeaderDragDrop = 0
    chkCheckBoxes = 0
    chkBoldFont = 0
    chkItalicFont = 0
    chkUnderlineFont = 0
    chkStrikeOutFont = 0
    chkBitmapsOnAllColumns = 0
    
    'Abilita l'opzione che pone a destra l'icona sull'intestazione
    chkImgPosition = 1
    
    'Imposta la posizione delle icone sulle intestazioni,
    'in questo caso a destra
    imgPosition = HDF_BITMAP_ON_RIGHT
    
End Sub

Private Sub chkFlatHeaders_Click()
    'Questa funzione definisce l'aspetto delle intestazioni
    'E' possibile scegliere fra piatto e 3D
    'Quando le intestazioni sono piatte non reagiscono al
    'click del mouse. Ideale per evitare che l'utente attivi
    'l'ordinamento di una colonna.
    
    Dim lStyle As Long
    Dim hHeader As Long
    
    'Ottiene l'handle dell'intestazione della listview
    hHeader = SendMessageLong(ListView1.hWnd, LVM_GETHEADER, 0, ByVal 0&)
   
    'Ottiene lo stile corrente per l'intestazione
    lStyle = GetWindowLong(hHeader, GWL_STYLE)
   
    'Modifica lo stile invertendo (da qui l'uso dello XOR) il flag HDS_BUTTONS
    'In questo modo ogni volta che si fa click, se è attivo lo disattiva
    'mentre se è disattivato lo attiva.
    'Se si desidera attivarlo e basta, utilizzare l'OR invece dello XOR
    lStyle = lStyle Xor HDS_BUTTONS
   
    'Imposta il nuovo stile e ridisegna la listview
    If lStyle Then
        Call SetWindowLong(hHeader, GWL_STYLE, lStyle)
        Call SetWindowPos(ListView1.hWnd, Me.hWnd, 0, 0, 0, 0, SWP_FLAGS)
    End If
End Sub

Private Sub chkGridLines_Click()
    'Questa funzione imposta la visualizzazione delle linee
    'di griglia normalmente non visibili nella ListView.
    
    Dim lStyle As Long
    
    'Ottiene lo stile esteso corrente
    lStyle = SendMessageLong(ListView1.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    
    'Modifica lo stile esteso invertendo il flag LVS_EX_GRIDLINES
    'In questo modo ogni volta che si fa click, se è attivo lo disattiva
    'mentre se è disattivato lo attiva.
    'Se si desidera attivarlo e basta, utilizzare l'OR invece dello XOR
    lStyle = lStyle Xor LVS_EX_GRIDLINES
    
    'Imposta il nuovo stile esteso
    Call SendMessageLong(ListView1.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, lStyle)

    'Per far ridisegnare l'intestazione ed eliminare l'inestetismo
    'dovuto al cambio di stile bisogna ricorrere ad un piccolo trucco...
    'Basta modificare la dimensione di una colonna
    'IMPOSTANDOLA ALLA STESSA DIMENSIONE DI PRIMA (?!?!?)
    'Se volete vedere cosa succede senza questo trucco, provate a
    'togliere la riga seguente...
    ListView1.ColumnHeaders.Item(1).Width = ListView1.ColumnHeaders.Item(1).Width
End Sub

Private Sub chkFullRowSelection_Click()
    'Questa funzione definisce il metodo di selezione delle
    'righe. Normalmente la ListView seleziona solo il main item
    'di ogni riga. E' possibile farle selezionare tutta la riga.
    
    Dim lStyle As Long
    
    'Ottiene lo stile esteso corrente
    lStyle = SendMessageLong(ListView1.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    
    'Modifica lo stile esteso invertendo il flag LVS_EX_FULLROWSELECT
    'In questo modo ogni volta che si fa click, se è attivo lo disattiva
    'mentre se è disattivato lo attiva.
    'Se si desidera attivarlo e basta, utilizzare l'OR invece dello XOR
    lStyle = lStyle Xor LVS_EX_FULLROWSELECT
    
    'Imposta il nuovo stile esteso
    Call SendMessageLong(ListView1.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, lStyle)
End Sub

Private Sub chkHeaderHotTrack_Click()
    'Questa funzione fa sì che le intestazioni cambino
    'colore al passaggio del mouse (funzionalità Hot Track)
    
    Dim hHeader As Long
    Dim lStyle As Long
    
    'Ottiene l'handle dell'intestazione della listview
    hHeader = SendMessageLong(ListView1.hWnd, LVM_GETHEADER, 0, 0)
    
    'Ottiene lo stile attuale dell'intestazione
    lStyle = GetWindowLong(hHeader, GWL_STYLE)
    
    'Modifica lo stile invertendo il flag HDS_HOTTRACK
    'In questo modo ogni volta che si fa click, se è attivo lo disattiva
    'mentre se è disattivato lo attiva.
    'Se si desidera attivarlo e basta, utilizzare l'OR invece dello XOR
    lStyle = lStyle Xor HDS_HOTTRACK
    
    'Imposta lo stile delle intestazioni
    Call SetWindowLong(hHeader, GWL_STYLE, lStyle)
End Sub
   
Private Sub chkHeaderDragDrop_Click()
    'Questa funzione imposta la caratteristica che permette
    'lo spostamento (drag & drop) delle colonne.
    
    Dim lStyle As Long
    
    'Ottiene lo stile esteso corrente
    lStyle = SendMessageLong(ListView1.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    
    'Modifica lo stile esteso invertendo il flag LVS_EX_HEADERDRAGDROP
    'In questo modo ogni volta che si fa click, se è attivo lo disattiva
    'mentre se è disattivato lo attiva.
    'Se si desidera attivarlo e basta, utilizzare l'OR invece dello XOR
    lStyle = lStyle Xor LVS_EX_HEADERDRAGDROP
    
    'Imposta il nuovo stile esteso
    Call SendMessageLong(ListView1.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, lStyle)
End Sub

Private Sub chkTrackSelect_Click()
    'Questa funzione fa sì che non sia necessario fare clic su un item
    'per selezionarlo. E' sufficiente passarci sopra e restarci per un po'.
    
    Dim lStyle As Long
    
    'Ottiene lo stile esteso corrente
    lStyle = SendMessageLong(ListView1.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    
    'Modifica lo stile esteso invertendo il flag LVS_EX_TRACKSELECT
    'In questo modo ogni volta che si fa click, se è attivo lo disattiva
    'mentre se è disattivato lo attiva.
    'Se si desidera attivarlo e basta, utilizzare l'OR invece dello XOR
    lStyle = lStyle Xor LVS_EX_TRACKSELECT
    
    'Imposta il nuovo stile esteso
    Call SendMessageLong(ListView1.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, lStyle)
End Sub

Private Sub chkCheckBoxes_Click()
    Dim lStyle As Long
    
    'Ottiene lo stile esteso corrente
    lStyle = SendMessageLong(ListView1.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    
    lStyle = lStyle Xor LVS_EX_CHECKBOXES
    
    'Imposta il nuovo stile esteso
    Call SendMessageLong(ListView1.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, lStyle)
End Sub

Private Sub chkImgPosition_Click()
    'Questa funzione imposta la posizione dell'icona
    'sull'intestazione (a destra o a sinistra)
    
    If chkImgPosition Then
         imgPosition = HDF_BITMAP_ON_RIGHT
    Else
         imgPosition = HDF_BITMAP
    End If
End Sub

Private Sub cmdAddHeaderPics_Click()
    'Mostra l'icona sull'intestazione della colonna specificata.
    'Il formato è:
    'SubName columnNo, imagelist iconNo, justification, image flag
    ShowHeaderIcon 1, 0, imgPosition, HDF_IMAGE
    
    'Disabilita la checkbox della posizione immagine
    chkImgPosition.Enabled = False
End Sub

Public Sub ShowHeaderIcon(colNo As Long, imgIconNo As Long, _
                          justify As Long, showImage As Long)

    'Questa funzione mostra l'icona sull'intestazione
    
    Dim r As Long
    Dim hHeader As Long
    Dim HD As HD_ITEM
    
    'Ottiene l'handle dell'intestazione della listview
    hHeader = SendMessageLong(ListView1.hWnd, LVM_GETHEADER, 0, 0)
    
    'Imposta i membri della struttura HD_ITEM
    With HD
        .mask = HDI_IMAGE Or HDI_FORMAT
        .fmt = HDF_LEFT Or HDF_STRING Or justify Or showImage
        .pszText = ListView1.ColumnHeaders(colNo + 1).Text
        If showImage Then
            .iImage = imgIconNo
        End If
    End With
    
    'Modifica l'intestazione
    r = SendMessageAny(hHeader, HDM_SETITEM, colNo, HD)
End Sub

Private Sub cmdCenteredImage_Click()
    'Questa funzione fa si che venga visualizzata un'immagine
    'al centro della ListView.
    'L'immagine viene letta dalla posizione specificata, per
    'cui occorre verificare che il file sia al suo posto...
    
    Dim BKIMG As LVBKIMAGE
    
    'Prepara la struttura di dati backimage
    With BKIMG
        'Questo flag dice al programma di utilizzare il nome
        'del file specificato nel membro pszImage
        .uFlags = LVBKIF_SOURCE_URL
        
        'il file di immagine da utilizzare
        .pszImage = "c:\windows\cerchi.bmp"
        
        'spostamento orizzontale e verticale dell'immagine
        'in percentuale. Un'impostazione di 0 allinea l'immagine
        'in alto a sinistra; 100 in basso a destra; 50 in mezzo
        .xOffsetPercent = 50
        .yOffsetPercent = 50
    End With
    
    'Imposta lo sfondo del testo a nessuno (trasparente)
    'in modo da far si che l'immagine si veda attraverso.
    'E' anche possibile specificare un valore RGB() per
    'colorare lo sfondo del testo.
    Call SendMessageLong(ListView1.hWnd, LVM_SETTEXTBKCOLOR, 0&, CLR_NONE)
    
    'Imposta l'immagine nella ListView
    Call SendMessageAny(ListView1.hWnd, LVM_SETBKIMAGE, 0, BKIMG)
End Sub

Private Sub cmdClear_Click()
    'Questa funzione toglie le icone dalle intestazioni

    Dim r As Long
    Dim colNo As Long
    Dim hHeader As Long
    Dim HD As HD_ITEM
    
    'Ottiene l'handle dell'intestazione della listview
    hHeader = SendMessageLong(ListView1.hWnd, LVM_GETHEADER, 0, 0)
    For colNo = 0 To ListView1.ColumnHeaders.Count - 1
        
        'Imposta i membri della struttura HD_ITEM
        With HD
            .mask = HDI_FORMAT
            .fmt = HDF_LEFT Or HDF_STRING
            .pszText = ListView1.ColumnHeaders(colNo + 1).Text
        End With
        
        'Modifica l'intestazione
        Call SendMessageAny(hHeader, HDM_SETITEM, colNo, HD)
    Next colNo
    
    'Riabilita la checkbox della posizione immagine
    chkImgPosition.Enabled = True
End Sub

Private Sub cmdNoDataImage_Click()
    'Questa funzione imposta l'immagine di sfondo
    'della ListView quando questa non contiene dati
    
    Dim BKIMG As LVBKIMAGE
    
    'In questo esempio eliminiamo appositamente tutte le righe
    'dalla ListView
    ListView1.ListItems.Clear
    
    With BKIMG
        If ListView1.ListItems.Count > 0 Then
            'Elimina qualsiasi immagine visualizzata
            .uFlags = LVBKIF_SOURCE_NONE
        Else
            'Imposta l'immagine di sfondo
            .uFlags = LVBKIF_SOURCE_URL
            .pszImage = "c:\windows\nuvole.bmp"
            .xOffsetPercent = 3
            .yOffsetPercent = 3
        End If
    End With
    
    'Imposta lo sfondo del testo a nessuno (trasparente)
    'in modo da far si che l'immagine si veda attraverso.
    'E' anche possibile specificare un valore RGB() per
    'colorare lo sfondo del testo.
    Call SendMessageLong(ListView1.hWnd, LVM_SETTEXTBKCOLOR, 0&, CLR_NONE)
    
    'Imposta l'immagine nella ListView
    Call SendMessageAny(ListView1.hWnd, LVM_SETBKIMAGE, 0&, BKIMG)
End Sub

Private Sub cmdSelectedRows_Click()
    'Questa funzione ha lo scopo di illustrare come
    'sia possibile capire quali items abbiano il segno
    'di spunta nella checkbox (se questa è stata abilitata)
    
    Dim i As Long
    Dim r As Long
    Dim lv As LVITEM
    Dim b As String
    b = "Le seguenti righe della ListView hanno il segno di spunta (0-based):" & vbCrLf & vbCrLf
    
    'cicla tutti gli items controllando il loro stato
    For i = 0 To ListView1.ListItems.Count - 1
        'Usa il messaggio LVM_GETITEMSTATE per leggere lo stato dell'item
        r = SendMessageLong(ListView1.hWnd, LVM_GETITEMSTATE, i, LVIS_STATEIMAGEMASK)
        'Quando un item ha il segno di spunta, la chiamata
        'LVM_GETITEMSTATE ritorna 8192 (&H2000&).
        If r And &H2000& Then
            'Ha il segno di spunta quindi impostiamo i membri della
            'struttura LVITEM
            With lv
                .cchTextMax = 255
                .pszText = Space$(255)
            End With
            
            'Ritroviamo il valore (testo) dell'item spuntato
            Call SendMessageAny(ListView1.hWnd, LVM_GETITEMTEXT, i, lv)
            
            b = b & "item " & CStr(i) & "  ( " & _
            Left$(lv.pszText, InStr(lv.pszText, Chr$(0)) - 1) & " )" & vbCrLf
      End If
      Next
      If b > "" Then MsgBox b
End Sub

Private Sub cmdTiledImage_Click()
   'Questa funzione mostra un'immagine ripetuta (affiancata)
   'all'interno della ListView.
   
   Dim BKIMG As LVBKIMAGE
   
   'Prepara la struttura backimage
   With BKIMG
      'Imposta il flag in modo da far si che usi l'immagine
      'specificata nel membro pszImage e fa si che l'immagine
      'sia affiancata
      .uFlags = LVBKIF_SOURCE_URL Or LVBKIF_STYLE_TILE
      'Specifica il file da utilizzare
      .pszImage = "c:\windows\cerchi.bmp"
    End With
    
    'Imposta lo sfondo del testo a nessuno (trasparente)
    'in modo da far si che l'immagine si veda attraverso.
    'E' anche possibile specificare un valore RGB() per
    'colorare lo sfondo del testo.
    Call SendMessageLong(ListView1.hWnd, LVM_SETTEXTBKCOLOR, 0&, CLR_NONE)
    
    'Imposta l'immagine nella ListView
    Call SendMessageAny(ListView1.hWnd, LVM_SETBKIMAGE, 0&, BKIMG)
End Sub

Private Sub cmdNoImage_Click()
    'Questa funzione toglie qualsiasi immagine dalla ListView

    Dim BKIMG As LVBKIMAGE
    
    'Prepara la struttura backimage
    With BKIMG
        'Imposta il flag in modo da non visualizzare
        'alcuna immagine
        .uFlags = LVBKIF_SOURCE_NONE
    End With
    
    'Imposta lo sfondo del testo a nessuno (trasparente)
    'in modo da far si che l'immagine si veda attraverso.
    'E' anche possibile specificare un valore RGB() per
    'colorare lo sfondo del testo.
    Call SendMessageLong(ListView1.hWnd, LVM_SETTEXTBKCOLOR, 0&, CLR_NONE)
    
    'Imposta l'immagine nella ListView
    Call SendMessageAny(ListView1.hWnd, LVM_SETBKIMAGE, 0, BKIMG)
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    'Questa funzione di evento risponde al click sull'intestazione
    'di ciascuna colonna mostrando l'icona relativa all'ordinamento
    'prescelto (freccia in su o in giù)

    Dim i As Long
    Static sOrder
    sOrder = Not sOrder
   
    'Usa l'ordinamento standard per ordinare gli items
    ListView1.SortKey = ColumnHeader.Index - 1
    ListView1.SortOrder = Abs(sOrder)
    ListView1.Sorted = True
    
    'Toglie l'immagine dall'intestazione della colonna
    'che non è più ordinata e la mette sull'intestazione
    'della colonna su cui l'utente ha fatto click.
    For i = 0 To 2
        If i = ListView1.SortKey Then
            ShowHeaderIcon ListView1.SortKey, ListView1.SortOrder, imgPosition, HDF_IMAGE
        Else
            ShowHeaderIcon i, 0, 0, 0
        End If
    Next
End Sub

Private Function InitComctl32(dwFlags As Long) As Boolean
    'Inizializza i controlli ComCtl32
    
    Dim icc As tagINITCOMMONCONTROLSEX
    
    On Error GoTo Err_OldVersion
    icc.dwSize = Len(icc)
    icc.dwICC = dwFlags
    
    'Visual Basic genera l'errore 453 "Specified 'DLL function not found"
    'se non è installata la nuova versione dei controlli e non è quindi
    'possibile trovare il nome della funzione.
    'Se siamo sfortunati riusciremo almeno a caricare la vecchia versione
    'nel codice seguente.
    
    InitComctl32 = InitCommonControlsEx(icc)
  
    Exit Function

Err_OldVersion:
    InitCommonControls

End Function

Private Sub cmdGetColumns_Click()
    'Questa funzione mostra la posizione attuale e quella originale
    'delle colonne (se queste sono state spostate con il drag & drop)

    Dim i As Long
    Dim r As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim totalCols As Long
    Dim Msg As String
    Dim tmp As String
    Dim LVC As LVCOLUMN
  
    'TotalCols è il totale delle colonne (1-based) richiesto
    'per la chiamata alle API.
    'LastCol è il numero (0-based) di colonne nella ListView,
    totalCols = ListView1.ColumnHeaders.Count
    firstCol = 0
    lastCol = totalCols - 1
      
    'Per ottenere l'ordine delle colonne occorre passare alle API
    'una matrice.
    'Al ritorno la matrice verrà riempita con l'indice delle colonne.
    ReDim posArray(firstCol To lastCol) As Long

    Call SendMessageAny(ListView1.hWnd, _
                        LVM_GETCOLUMNORDERARRAY, _
                        totalCols, _
                        posArray(firstCol))
       
    For i = firstCol To lastCol
        tmp = Space$(32)
          
        With LVC
            .mask = LVCF_TEXT
            .pszText = tmp
            .cchTextMax = Len(tmp)
        End With
          
        Call SendMessageAny(ListView1.hWnd, LVM_GETCOLUMN, posArray(i), LVC)

        tmp = Left$(LVC.pszText, InStr(LVC.pszText, Chr$(0)) - 1)

        Msg = Msg & vbTab & tmp & vbTab & vbTab & posArray(i) & vbTab & vbCrLf
    Next

    MsgBox "    Ordine corrente / Indice originale (0 based): " & vbCrLf & vbCrLf & Msg
End Sub

Private Sub SetHeaderFontStyle()
    'Questa funzione imposta lo stile del carattere delle intestazioni

    Dim LF As LOGFONT
    Dim r As Long
    Dim hCurrFont As Long
    Dim hOldFont As Long
    Dim hHeader As Long
    
    'Ottiene l'handle dell'intestazione
    hHeader = SendMessageLong(ListView1.hWnd, LVM_GETHEADER, 0, 0)
    
    'Ottiene l'handle del font utilizzato nell'intestazione
    hCurrFont = SendMessageLong(hHeader, WM_GETFONT, 0, 0)
   
    'Ottiene i dettagli LOGFONT del font utilizzato
    r = GetObject(hCurrFont, Len(LF), LF)
    If r > 0 Then
        'Imposta gli attributi del font in base alle checkboxes relative
        'sulla maschera
        
        If chkBoldFont = 1 Then
            LF.lfWeight = FW_BOLD
        Else
            LF.lfWeight = FW_NORMAL
        End If
        
        LF.lfItalic = chkItalicFont
            
        LF.lfUnderline = chkUnderlineFont

        LF.lfStrikeOut = chkStrikeOutFont
        
        'Elimina il font precedente
        r = DeleteObject(hHeaderFont)
        
        'Crea il nuovo font
        hHeaderFont = CreateFontIndirect(LF)
        
        'Seleziona il nuovo font
        hOldFont = SelectObject(hHeader, hHeaderFont)
        
        'Informa la ListView che il font è cambiato
        r = SendMessageLong(hHeader, WM_SETFONT, hHeaderFont, True)
    End If
End Sub

Private Sub chkBoldFont_Click()
    SetHeaderFontStyle
End Sub

Private Sub chkItalicFont_Click()
    SetHeaderFontStyle
End Sub

Private Sub chkUnderlineFont_Click()
    SetHeaderFontStyle
End Sub

Private Sub chkStrikeOutFont_Click()
    SetHeaderFontStyle
End Sub

Private Sub cmdAdaptColumn2_Click()
    'Adatta la larghezza della colonna 2 alla larghezza
    'del testo dell'intestazione
    Dim lColumnIndex As Long
    
    lColumnIndex = 1
    
    'Se si include l'intestazione del ridimensionamento, allora
    'l'ultima colonna verrà automaticamente ridimensionata in modo
    'da riempire lo spazio restante della ListView.
    With ListView1
        'Verifica che la ListView sia nella visualizzazione Report
        If .View = lvwReport Then
            Call SendMessage(.hWnd, LVM_SETCOLUMNWIDTH, lColumnIndex, ByVal LVSCW_AUTOSIZE_USEHEADER)
        End If
    End With
End Sub

Private Sub cmdAdaptColumn2NoHeader_Click()
    'Adatta la larghezza della colonna 2 alla larghezza
    'del suo contenuto (non del testo dell'intestazione)
    
    Dim lColumnIndex As Long
   
    lColumnIndex = 1
   
    'Se si include l'intestazione del ridimensionamento, allora
    'l'ultima colonna verrà automaticamente ridimensionata in modo
    'da riempire lo spazio restante della ListView.
    With ListView1
        'Verifica che la ListView sia nella visualizzazione Report
        If .View = lvwReport Then
            Call SendMessage(.hWnd, LVM_SETCOLUMNWIDTH, lColumnIndex, ByVal LVSCW_AUTOSIZE)
        End If
    End With
End Sub

Private Sub chkBitmapsOnAllColumns_Click()
    Dim lStyle As Long
    
    lStyle = SendMessageLong(ListView1.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    lStyle = lStyle Xor LVS_EX_SUBITEMIMAGES
    
    'Imposta il nuovo stile esteso
    Call SendMessageLong(ListView1.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, lStyle)
    
    ListView1.Refresh
End Sub

Private Sub cmdChangeRowTextColor_Click()
    'Questa funzione cambia il colore del testo delle righe

    Dim lTextColor As Long
    
    'Ottiene il colore attuale del testo
    lTextColor = SendMessageLong(ListView1.hWnd, LVM_GETTEXTCOLOR, 0, 0)

    'Se il colore è rosso lo reimposta a nero, altrimenti se è
    'nero lo imposta a rosso
    If lTextColor = vbRed Then
        Call SendMessageLong(ListView1.hWnd, LVM_SETTEXTCOLOR, 0, vbBlack)
    Else
        Call SendMessageLong(ListView1.hWnd, LVM_SETTEXTCOLOR, 0, vbRed)
    End If
    
    'Ridisegna la ListView
    ListView1.Refresh
End Sub

Private Sub cmdChangeRowBackColor_Click()
    'Questa funzione cambia il colore di sfondo delle righe

    Dim lTextColor As Long
    
    'Ottiene il colore corrente del testo
    lTextColor = SendMessageLong(ListView1.hWnd, LVM_GETTEXTBKCOLOR, 0, 0)
    
    'Se è il colore da me scelto, lo reimposta a nessun colore
    'altrimenti lo imposta al colore da me scelto
    If lTextColor = 13565853 Then
        Call SendMessageLong(ListView1.hWnd, LVM_SETTEXTBKCOLOR, 0, CLR_NONE)
    Else
        Call SendMessageLong(ListView1.hWnd, LVM_SETTEXTBKCOLOR, 0, 13565853)
    End If
    
    'Ridisegna la ListView
    ListView1.Refresh
    
    'Se qualcuno vuole cimentarsi nel colorare solo alcuni items...
    'Call SendMessageLong(ListView1.hWnd, LVM_REDRAWITEMS, 1, 1)
    'Call SendMessageLong(ListView1.hWnd, LVM_REDRAWITEMS, 3, 3)
    'Call SendMessageLong(ListView1.hWnd, LVM_REDRAWITEMS, 5, 5)
    'Call SendMessageLong(ListView1.hWnd, LVM_REDRAWITEMS, 7, 7)
End Sub

Private Sub cmdIndentItem_Click()
    'Questa funzione indenta la riga corrente.
    'Utile per simulare una treeview (avete presente Outlook Express nelle news ?)

    Dim tLV As LVITEM
    Dim lR As Long
    Dim lI As Long

    tLV.mask = LVIF_INDENT
    'L'indice della ListView non corrisponde all'indice delle API
    'se è stao attuato un ordinamento personalizzato
    tLV.iItem = ListView1.SelectedItem.Index - 1
    
    lR = LVNI_ALL Or LVNI_SELECTED
    lI = SendMessageLong(ListView1.hWnd, LVM_GETNEXTITEM, -1, lR)
    tLV.iItem = lI
    
    If (SendMessage(ListView1.hWnd, LVM_GETITEM, 0, tLV) <> 0) Then
        '1 Unità di indentazione = Larghezza dell'icona
        
        tLV.iIndent = tLV.iIndent + 1
        SendMessage ListView1.hWnd, LVM_SETITEM, 0, tLV
    End If
End Sub
