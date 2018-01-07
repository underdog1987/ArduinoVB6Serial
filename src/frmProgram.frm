VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgram 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comunicación Serial y Arduino"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6570
   Icon            =   "frmProgram.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5280
      Top             =   0
   End
   Begin VB.Frame Frame4 
      Caption         =   "Raw"
      Height          =   855
      Left            =   240
      TabIndex        =   15
      Top             =   4920
      Width           =   6135
      Begin VB.TextBox txRaw 
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   5895
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Potenciómetro"
      Height          =   1095
      Left            =   240
      TabIndex        =   11
      Top             =   3600
      Width           =   4935
      Begin MSComctlLib.ProgressBar barrita 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Max             =   254
      End
      Begin VB.TextBox txPotenciometro 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pulsador"
      Height          =   2295
      Left            =   4320
      TabIndex        =   10
      Top             =   1080
      Width           =   1935
      Begin VB.Shape shCirculoPulsador 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   1695
         Left            =   120
         Shape           =   3  'Circle
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.ComboBox cmbPuerto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2760
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Led RGB"
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   3615
      Begin VB.HScrollBar sldBlue 
         Height          =   255
         LargeChange     =   10
         Left            =   720
         Max             =   254
         Min             =   3
         TabIndex        =   6
         Top             =   1680
         Value           =   3
         Width           =   1935
      End
      Begin VB.HScrollBar sldGreen 
         Height          =   255
         LargeChange     =   10
         Left            =   720
         Max             =   254
         Min             =   3
         TabIndex        =   5
         Top             =   1080
         Value           =   3
         Width           =   1935
      End
      Begin VB.HScrollBar sldRed 
         Height          =   255
         LargeChange     =   10
         Left            =   720
         Max             =   254
         Min             =   3
         TabIndex        =   2
         Top             =   480
         Value           =   3
         Width           =   1935
      End
      Begin VB.PictureBox pColor 
         FillStyle       =   0  'Solid
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1515
         ScaleWidth      =   435
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblBlue 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblGreen 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblRed 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Arduino en puerto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "frmProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sendColor As Byte
Private Sub cmbPuerto_Change()
    cmbPuerto_Click
End Sub

Private Sub cmbPuerto_Click()
    Dim iPort As Integer, a As Integer
    If MSComm1.PortOpen Then
        MSComm1.PortOpen = False
        For a = 1 To 1024: DoEvents: DoEvents: Next
    End If
    iPort = CInt(Right$(cmbPuerto.Text, Len(cmbPuerto.Text) - 3))
    
    MSComm1.CommPort = iPort
    MSComm1.PortOpen = True
    colorInicial
    
End Sub

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub Form_Load()
    Dim x As Integer
    For x = 1 To 255
        cmbPuerto.AddItem "COM" & x
    Next
    cmbPuerto.Text = cmbPuerto.List(0)
    
    shCirculoPulsador.FillStyle = 1
    'Color inicial
    colorInicial
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim a As Integer
    If MSComm1.PortOpen Then
        MSComm1.PortOpen = False
        For a = 1 To 1024: DoEvents: DoEvents: Next
    End If
End Sub

Private Sub sldBlue_Change()
    sendColor = 2
    setColors sldRed.Value, sldGreen.Value, sldBlue.Value
End Sub

Private Sub sldGreen_Change()
    sendColor = 1
    setColors sldRed.Value, sldGreen.Value, sldBlue.Value
End Sub

Private Sub sldRed_Change()
    sendColor = 0
    setColors sldRed.Value, sldGreen.Value, sldBlue.Value
End Sub

Private Sub Timer1_Timer()
Dim sRecibido As String ' Byte recibido como string
Dim bRecibido As Byte ' byte recibido

If MSComm1.PortOpen Then
    'r = MSComm1.Input
    sRecibido = MSComm1.Input
    If sRecibido = "" Then
        bRecibido = 0
    Else
        bRecibido = Asc(sRecibido)
    End If
    If bRecibido = 255 Then
        shCirculoPulsador.FillStyle = 0
    Else
        barrita.Value = bRecibido
        txPotenciometro.Text = bRecibido
        shCirculoPulsador.FillStyle = 1
    End If
    txRaw.Text = sRecibido
    
    
End If
End Sub

Private Sub setColors(r As Integer, g As Integer, b As Integer)
    Dim p As Integer
    MSComm1.Output = Chr$(sendColor)
    For p = 0 To 16384
        DoEvents
    Next
    lblRed.Caption = r
    lblGreen.Caption = g
    lblBlue.Caption = b
    pColor.BackColor = RGB(r, g, b)
    Select Case sendColor
        Case Is = 0
            MSComm1.Output = Chr$(r)
        Case Is = 1
            MSComm1.Output = Chr$(g)
        Case Is = 2
            MSComm1.Output = Chr$(b)
    End Select
End Sub

Private Sub colorInicial()
    sldRed.Value = 3
    sldGreen.Value = 3
    sldBlue.Value = 3
    sendColor = 0
    setColors 3, 3, 3
    sendColor = 1
    setColors 3, 3, 3
    sendColor = 2
    setColors 3, 3, 3
End Sub

