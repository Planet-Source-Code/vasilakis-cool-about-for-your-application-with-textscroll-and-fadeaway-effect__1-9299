VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About!"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2655
      Left            =   600
      ScaleHeight     =   2595
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   840
      Width           =   4455
      Begin VB.Label lblAn 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "A Cool About Effect by Vasilis Sagonas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   1080
         Visible         =   0   'False
         Width           =   4410
      End
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   5160
      Top             =   3480
   End
   Begin VB.Label lblinfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I Hope you like it! If you do, please vote for me!"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

'Load as many lines as you want!

'********************
  LinesNum = 6
'********************


'*************************************
'Initialize and load Labels and Timers
For i = 1 To LinesNum
    Load lblAn(i)
    Load tmr(i)
    lblAn(i).Top = lblAn(i).Top
    lblAn(i).Left = lblAn(i).Left
    lblAn(i).Height = lblAn(i).Height
    lblAn(i).Width = lblAn(i).Width
    lblAn(i).Visible = True
    
    'If its too fast, change here! If its too slow, try changing the speed from the
    'source of the timer.
    tmr(i).Interval = 1
    
    tmr(i).Enabled = False
Next i



'*************************************
'Using the arrays, add your lines!

lblAn(1).Caption = "About YOU! By Vasilis Sagonas"

'If you want more than one lines per scroll, add a vbCrLf like this:
lblAn(2).Caption = "A Nice About TEXT Scroller with a cool Effect!" & vbCrLf & "Only A Few Lines Of Code!"

lblAn(3).Caption = "Programmed by me!"

lblAn(4).Caption = "Blah blah blah BLAH!"

lblAn(5).Caption = "Bling Blung Haha!"

lblAn(6).Caption = "" 'Leave the last one blank if you want some seconds before restart the about text.
'*************************************



'*************************************
'Initialize positions

For i = 1 To lblAn.Count - 1
    lblAn(i).FontSize = 9
    lblAn(i).AutoSize = True
    lblAn(i).Alignment = vbCenter
    lblAn(i).Top = Picture1.Height + 10
Next i
'*************************************



'*************************************
'Start the initialization timer! Everything now depends on the timers!
tmr(0).Enabled = True
'*************************************


Show 'Show Form


'*************************************
'Create this beautiful Blue to Black Background

LineBackGround Picture1

'*************************************
End Sub


Private Sub LineBackGround(ctl As PictureBox)
On Error Resume Next

'This line back ground code is taken from the Setup1.VBP (Visual Basic Setup)
'I made some changes...!
    
    Const intBLUESTART% = 255
    Const intBLUEEND% = 0
    Const intBANDHEIGHT% = 50

    Dim sngBlueCur As Single
    Dim sngBlueStep As Single
    Dim intFormHeight As Integer
    Dim intFormWidth As Integer
    Dim intY As Integer

    intFormHeight = ctl.ScaleHeight
    intFormWidth = ctl.ScaleWidth
    
    sngBlueStep = intBANDHEIGHT * (intBLUEEND - intBLUESTART) / intFormHeight
    sngBlueCur = intBLUESTART

    For intY = 0 To intFormHeight Step intBANDHEIGHT
        ctl.Line (-1, intY - 1)-(intFormWidth, intY + intBANDHEIGHT), RGB(0, 0, sngBlueCur), BF
        sngBlueCur = sngBlueCur + sngBlueStep
    Next intY
End Sub


Private Sub tmr_Timer(Index As Integer)
On Error Resume Next
If Index = 0 Then
    
    tmr(0).Enabled = False
    tmr(1).Enabled = True
Else
    
    '*****************************************
    'This is the speed which the labels goes.
    'Default is 5. Change it to make it faster or slower...
    'For Slower, I recommend you to change the timer's Interval.
    lblAn(Index).Top = lblAn(Index).Top - 5
    '*****************************************
    
    FS = (CInt(lblAn(Index).Top) / 100) * 3
    
    
    '************************************************************************
    'Wanna change the color effect? Play with this!
    
    'If you want you can do incredible things with colors.
    'Add onother one variable like FCol (Fcol = foreground color for me)
    'and add it below at the RGB(..) Function and you'll have
    'awesome results!
    
    FCol = CInt((2000 / Picture1.Height) * CInt(lblAn(Index).Top))
    
    ' * FCol2 = CInt((1000 / Picture1.Height) * CInt(lblAn(Index).Top))
    ' * FCol3 = CInt((100 / Picture1.Height) * CInt(lblAn(Index).Top))
    
   
    
    lblAn(Index).ForeColor = RGB(FCol, FCol, FCol) 'White to Black
    
    'This one uses the FCol2 and FCol3... Try it!
    'lblAn(Index).ForeColor = RGB(FCol, FCol2, FCol3)
    '************************************************************************
    
    '************************************************************************
    'Just a check for bugs at the FontSize.
    If FS <= 0 Then FS = 1
    If FS > 9 Then FS = 9
    '************************************************************************
    
    
    '************************************************************************
    'Check for the FontSize and the Fade Away Effect Speed
    
    'Just check if the FontSize is not already gaven (for speed waste)
    If Not lblAn(Index).FontSize = FS Then
        
        lblAn(Index).FontSize = FS
        
        'If you want to change the Fade Away Speed, Change the value below.
        'Default is 2.
        lblAn(Index).Top = lblAn(Index).Top - 2
        
    End If
    '************************************************************************
    
    
    
    '************************************************************************
    'Label is Over the head! It walked fine! Let's go to the next one
    'if there is!
    
    If lblAn(Index).Top <= 0 - lblAn(Index).Height - 10 Then 'Check if its over the PictureBox
        
        tmr(Index).Enabled = False 'Stop current Timer.
        
        tmr(Index + 1).Enabled = True 'Start Next Timer For next Label Scroll
        
        
        'Check if we are at the end of the text(labels).
        If Index + 1 = tmr.Count Then
            
            'If Yes, RESET Every labels Font/Top and Start from the beginning!
            For i = 1 To lblAn.Count
                lblAn(i).FontSize = 5
                lblAn(i).Top = Picture1.Height + 10
            Next i
            
            'Start The First Timer just for time waste.
            tmr(0).Enabled = True
        
        End If
    
    End If

End If

'That's it! It wasn't too hard, was it??
'If you like this code email me and
'if you really really like this code
'please vote at PSC! And I really want
'your comments! If you make anything crazy,
'please send it to me! I will apreciate it!

'I originally made this code for NetFLY! v1.5
'http://www.dawn.gr/netfly
'And i thought that it would be a great idea to
'share it with you! :-)

'Vasilis Sagonas
'Official Web Site: http://www.dawn.gr/
'Email: vasilis@lar.forthnet.gr/vsag@forthnet.gr/webmaster@dawn.gr

End Sub


