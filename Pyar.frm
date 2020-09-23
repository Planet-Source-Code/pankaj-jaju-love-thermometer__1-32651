VERSION 5.00
Begin VB.Form Pyar 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000CC&
   BorderStyle     =   0  'None
   Caption         =   " Love Thermometer"
   ClientHeight    =   6360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7710
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TextFirst 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000CC&
      Height          =   405
      Left            =   795
      MaxLength       =   40
      TabIndex        =   1
      Text            =   "ENTER FIRST NAME"
      ToolTipText     =   "Enter First Person's Name"
      Top             =   1530
      Width           =   2355
   End
   Begin VB.TextBox TextSecond 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000CC&
      Height          =   405
      Left            =   4560
      MaxLength       =   40
      TabIndex        =   2
      Text            =   "ENTER SECOND NAME"
      ToolTipText     =   "Enter Second's Person's Name"
      Top             =   1530
      Width           =   2355
   End
   Begin VB.CommandButton CmdCalculate 
      BackColor       =   &H0080FFFF&
      Caption         =   "Calculate Love "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3338
      MaskColor       =   &H000000CC&
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Click Me to calculate the Love Between both persons"
      Top             =   3165
      UseMaskColor    =   -1  'True
      Width           =   1035
   End
   Begin VB.PictureBox PicTitleBox 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   7710
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   7710
      Begin VB.Image ImgClose 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   7215
         Picture         =   "Pyar.frx":0000
         ToolTipText     =   "Close Me"
         Top             =   60
         Width           =   405
      End
      Begin VB.Image ImgMin 
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   6750
         Picture         =   "Pyar.frx":006F
         ToolTipText     =   "Minimize Me"
         Top             =   60
         Width           =   405
      End
   End
   Begin VB.Label LabelCaption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "created by:- PANKAJ JAJU"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   420
      Left            =   600
      TabIndex        =   5
      ToolTipText     =   "created by:- PANKAJ JAJU"
      Top             =   4125
      Width           =   6510
   End
   Begin VB.Label LabelInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000CC&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1500
      Left            =   600
      TabIndex        =   4
      Top             =   4560
      Width           =   6510
   End
   Begin VB.Shape ShapeInfo 
      BackColor       =   &H000000CC&
      BorderColor     =   &H00800000&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   2010
      Left            =   540
      Top             =   4110
      Width           =   6630
   End
   Begin VB.Line LineBorder3 
      BorderColor     =   &H00800000&
      BorderWidth     =   5
      X1              =   0
      X2              =   7710
      Y1              =   6310
      Y2              =   6300
   End
   Begin VB.Line LineBorder2 
      BorderColor     =   &H00800000&
      BorderWidth     =   5
      X1              =   7665
      X2              =   7665
      Y1              =   540
      Y2              =   6255
   End
   Begin VB.Line LineBorder1 
      BorderColor     =   &H00800000&
      BorderWidth     =   5
      X1              =   30
      X2              =   30
      Y1              =   525
      Y2              =   6240
   End
   Begin VB.Image ImgFirst 
      Appearance      =   0  'Flat
      Height          =   3300
      Left            =   105
      Picture         =   "Pyar.frx":00C6
      Top             =   765
      Width           =   3750
   End
   Begin VB.Image ImgSecond 
      Appearance      =   0  'Flat
      Height          =   3300
      Left            =   3855
      Picture         =   "Pyar.frx":0CEE
      Top             =   765
      Width           =   3750
   End
End
Attribute VB_Name = "Pyar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' (((((((( (((((((( (((    ((( (((  (((( (((((((( (((((((((
' )))  ))) )))  ))) ))))   ))) ))) ))))  )))  )))     )))
' (((((((( (((((((( ((( (( ((( (((((((   ((((((((     (((
' )))      )))))))) )))  ))))) ))) ))))  ))))))))     )))
' (((      (((  ((( (((   (((( (((   ((( (((  ((( ((  (((
' )))      )))  ))) )))    ))) (((   ((( (((  ((( )))))))
'
'«··´¯`·.¸_¸.·´¯`·.¸_¸.·´¯`·.¸_¸.·´¯`·.¸_¸.·´¯`·.¸_¸..·´¯`·.
'INDIA IS GREAT                                          .·´
'Pankaj Jaju                                             ´·.
'E-mail: pankaj_jaju@rediffmail.com                       .·´
'PLEASE SEND YOUR VALUABLE SUGGESTIONS                  ´·.
'«·´¯`·.¸_¸.·´¯`·.¸_¸.·´¯`·.¸_¸.·´¯`·.¸_¸.·´¯`·.¸_¸.·´¯`·..·´

'******************************************************************************************

Option Explicit
Dim XX As Single, YY As Single
Dim MOVEFORM As Boolean, CMDCAL As Boolean

Private Sub CmdCalculate_Click()
    If Len(Trim(TextFirst.Text)) = 0 Or Len(Trim(TextSecond.Text)) = 0 Then
        If Len(Trim(TextFirst.Text)) = 0 And Len(Trim(TextSecond.Text)) = 0 Then
            MsgBox "You Left Both the TextBoxes EMPTY", vbCritical, "Invalid Input"
            TextFirst.SetFocus
            Exit Sub
        End If
        If Len(Trim(TextFirst.Text)) = 0 Then
            MsgBox "You Left The First TextBox EMPTY", vbCritical, "Invalid Input"
            TextFirst.SetFocus
        ElseIf Len(Trim(TextSecond.Text)) = 0 Then
            MsgBox "You Left The Second TextBox EMPTY", vbCritical, "Invalid Input"
            TextSecond.SetFocus
        End If
        Exit Sub
    End If
    
    Dim FIRST As Byte, SECOND As Byte, COUNT As Byte
    Dim LOVE As Integer
    Dim LETTER As String * 1
    
    FIRST = Len(Trim(TextFirst.Text))
    SECOND = Len(Trim(TextSecond.Text))

    For COUNT = 1 To FIRST
        LETTER = Mid$(UCase(Trim(TextFirst.Text)), COUNT, 1)
        If LETTER = "I" Then LOVE = LOVE + 1
        If LETTER = "L" Then LOVE = LOVE + 2
        If LETTER = "O" Then LOVE = LOVE + 1
        If LETTER = "V" Then LOVE = LOVE + 3
        If LETTER = "E" Then LOVE = LOVE + 1
        If LETTER = "Y" Then LOVE = LOVE + 3
        If LETTER = "U" Then LOVE = LOVE + 3
    Next
    For COUNT = 1 To SECOND
        LETTER = Mid$(UCase(Trim(TextSecond.Text)), COUNT, 1)
        If LETTER = "I" Then LOVE = LOVE + 1
        If LETTER = "L" Then LOVE = LOVE + 2
        If LETTER = "O" Then LOVE = LOVE + 1
        If LETTER = "V" Then LOVE = LOVE + 3
        If LETTER = "E" Then LOVE = LOVE + 1
        If LETTER = "Y" Then LOVE = LOVE + 3
        If LETTER = "U" Then LOVE = LOVE + 3
    Next
    
    LOVE = (LOVE / (FIRST + SECOND)) * 100
    If LOVE > 99 Then LOVE = 100
    
    LabelInfo.Caption = ""
    LabelInfo.Font.Size = 20
    
    If LOVE <= 15 Then LabelInfo.Caption = "'BORN ENEMIES'"
    If LOVE > 15 And LOVE <= 30 Then LabelInfo.Caption = "'I HATE YOU'"
    If LOVE > 30 And LOVE <= 50 Then LabelInfo.Caption = "'YA I KNOW HIM'"
    If LOVE > 50 And LOVE <= 65 Then LabelInfo.Caption = "'WE ARE FRIENDS'"
    If LOVE > 65 And LOVE <= 80 Then LabelInfo.Caption = "'GETTING HOTTER'"
    If LOVE > 80 And LOVE <= 90 Then LabelInfo.Caption = "'WE R IN LOVE'"
    If LOVE > 90 And LOVE <= 95 Then LabelInfo.Caption = "'KISS ME NOW'"
    If LOVE > 95 Then LabelInfo.Caption = "'THAT'S INDIAN STYLE'"
    
    LabelInfo.Caption = LOVE & "°C" & vbCrLf & LabelInfo.Caption
    CMDCAL = True
End Sub

Private Sub CmdCalculate_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbAltMask And KeyCode = vbKeyF4 Then End
End Sub

Private Sub Form_Load()
    PicTitleBox.Print " Love Thermometer"
    LabelInfo.Caption = vbCrLf & "Hi Guys! This is my second VB Project after ImageViewer" & _
                        vbCrLf & "This programme simply calculates the Love % between 2 persons" & _
                        vbCrLf & "Please do find some time to RATE this code"
    CMDCAL = True
End Sub

Private Sub ImgClose_Click()
    End
End Sub

Private Sub ImgClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgClose.BorderStyle = 0
End Sub

Private Sub ImgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgClose.BorderStyle = 1
End Sub

Private Sub ImgMin_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub ImgMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgMin.BorderStyle = 0
End Sub

Private Sub ImgMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgMin.BorderStyle = 1
End Sub

Private Sub PicTitleBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    XX = X
    YY = Y
    MOVEFORM = True
End Sub

Private Sub PicTitleBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MOVEFORM = True Then
        Me.Left = Me.Left + (X - XX)
        Me.Top = Me.Top + (Y - YY)
    End If
End Sub

Private Sub PicTitleBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Left = Me.Left + (X - XX)
    Me.Top = Me.Top + (Y - YY)
    If MOVEFORM = True Then MOVEFORM = False: Me.Refresh: TextFirst.SetFocus
End Sub

Private Sub TextFirst_GotFocus()
    If CMDCAL = True Then
        LabelInfo.Caption = ""
        LabelInfo.Font.Size = 10
        LabelInfo.Caption = vbCrLf & "Hi Guys! This is my second VB Project after ImageViewer" & _
                            vbCrLf & "This programme simply calculates the Love % between 2 persons" & _
                            vbCrLf & "Please do find some time to RATE this code"
        TextSecond.Text = ""
        TextFirst.Text = ""
        CMDCAL = False
    End If
End Sub

Private Sub TextFirst_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbAltMask And KeyCode = vbKeyF4 Then End
End Sub

Private Sub TextSecond_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbAltMask And KeyCode = vbKeyF4 Then End
End Sub
