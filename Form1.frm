VERSION 5.00
Object = "{E7BC34A0-BA86-11CF-84B1-CBC2DA68BF6C}#1.0#0"; "NTSVC.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo cara mudah membuat NT Service"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5190
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAnimasi 
      Left            =   1920
      Top             =   1440
   End
   Begin NTService.NTService NTService1 
      Left            =   3480
      Top             =   960
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      ServiceName     =   "Simple"
      StartMode       =   3
   End
   Begin VB.PictureBox picPanel 
      Height          =   2055
      Left            =   120
      ScaleHeight     =   1995
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.Label lblAnimasi 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Service ...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
' http://coding4ever.wordpress.com                                         '
'***************************************************************************

Option Explicit

Dim ZigZagX     As Integer
Dim ZigZagY     As Integer

Dim stopService As Boolean

Private Sub initNTService()
    On Error GoTo ServiceError
    
    stopService = False
    
    With NTService1
        .DisplayName = "Coding4ever NT Service"
        .ServiceName = "coding4everNTService"
            
        lblAnimasi.Caption = .DisplayName & " Loading"
        
        'Install the service
        If Command$ = "/i" Then
        
            'enable interaction with desktop
            .Interactive = True
            .StartMode = svcStartAutomatic
            
            'Install the program as an NT service
            If .Install Then
                'Save the TimerInterval Parameter in the Registry
                .SaveSetting "Parameters", "TimerInterval", "45"
                
                MsgBox .DisplayName & ": installed successfully"
                
            Else
                MsgBox .DisplayName & ": failed to install"
            End If
            
            End
            
        'Remove the Service Registry Keys and uninstall the service
        ElseIf Command$ = "/u" Then
            If .Uninstall Then
                MsgBox .DisplayName & ": uninstalled successfully"
            Else
                MsgBox .DisplayName & ": failed to uninstall"
            End If
            
            End
            
        'Invalid parameter
        ElseIf Command$ <> "" Then
            MsgBox "Invalid Parameter"
            End
        End If
        
        'Retrive the stored value for the timer interval
        tmrAnimasi.Interval = CInt(.GetSetting("Parameters", "TimerInterval", "45"))
        
        'enable Pause/Continue. Must be set before StartService
        'is called or in design mode
        .ControlsAccepted = svcCtrlPauseContinue
        
        'connect service to Windows NT services controller
        .StartService
    End With
    
    Exit Sub
ServiceError:
    Call NTService1.LogEvent(svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description)
End Sub

Private Sub Form_Load()
    ZigZagX = 20
    ZigZagY = 20
    Call initNTService
End Sub

Private Sub NTService1_Continue(Success As Boolean)
    'Handle the continue service event
    On Error GoTo ServiceError
    
    tmrAnimasi.Enabled = True
    lblAnimasi.Caption = NTService1.DisplayName & " Running"
    Success = True
    
    NTService1.LogEvent svcEventInformation, svcMessageInfo, "Service continued"
    
    Exit Sub
ServiceError:
    NTService1.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub

Private Sub NTService1_Control(ByVal mEvent As Long)
    'Take control of the service events
    On Error GoTo ServiceError
    
    lblAnimasi.Caption = NTService1.DisplayName & " Control signal " & CStr([mEvent])
    
    Exit Sub
ServiceError:
    NTService1.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub

Private Sub NTService1_Pause(Success As Boolean)
    'Pause Event Request
    On Error GoTo ServiceError
    
    tmrAnimasi.Enabled = False
    lblAnimasi.Caption = NTService1.DisplayName & " Paused"
    NTService1.LogEvent svcEventError, svcMessageError, "Service paused"
    Success = True
    
    Exit Sub
ServiceError:
    NTService1.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub

Private Sub NTService1_Start(Success As Boolean)
    'Start Event Request
    On Error GoTo ServiceError
    
    lblAnimasi.Caption = NTService1.DisplayName & " Running"
    Success = True
    
    Exit Sub
ServiceError:
    NTService1.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub

Private Sub NTService1_Stop()
    'Stop and terminate the Service
    On Error GoTo ServiceError
    
    lblAnimasi.Caption = NTService1.DisplayName & " Stopped"
    stopService = True
    
    Unload Me
    
ServiceError:
    NTService1.LogEvent svcMessageError, svcEventError, "[" & Err.Number & "] " & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer) 'Unload the Service
    If Not stopService Then
        If MsgBox("Are you sure you want to unload the service?..." & vbCrLf & "the service will be stopped", vbQuestion + vbYesNo, "Stop Service") = vbYes Then
            NTService1.stopService
            lblAnimasi.Caption = NTService1.DisplayName & " Stopping"

            Cancel = True
            
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub tmrAnimasi_Timer()
    lblAnimasi.Move lblAnimasi.Left + ZigZagX, lblAnimasi.Top + ZigZagY
    If lblAnimasi.Left < picPanel.ScaleLeft Then
        ZigZagX = 20
    ElseIf lblAnimasi.Left + lblAnimasi.Width > picPanel.ScaleWidth + picPanel.ScaleLeft Then
        ZigZagX = -20
    ElseIf lblAnimasi.Top < picPanel.ScaleTop Then
        ZigZagY = 20
    ElseIf lblAnimasi.Top + lblAnimasi.Height > picPanel.ScaleHeight + picPanel.ScaleTop Then
        ZigZagY = -20
    End If
End Sub
