VERSION 5.00
Begin VB.Form Form_Params 
   Caption         =   "Params"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Thruster_Frame 
      Caption         =   "Thruster"
      Height          =   3735
      Left            =   3960
      TabIndex        =   52
      Top             =   3480
      Width           =   5055
      Begin VB.CommandButton Command5 
         Caption         =   "Help"
         Height          =   255
         Left            =   2040
         TabIndex        =   56
         Top             =   3360
         Width           =   855
      End
      Begin VB.Frame Thruster_T2_Frame 
         Caption         =   "T2"
         Height          =   3015
         Left            =   2520
         TabIndex        =   55
         Top             =   240
         Width           =   2415
         Begin VB.TextBox Thrtust_T2_Skew 
            Height          =   285
            Left            =   1200
            TabIndex        =   86
            Top             =   1680
            Width           =   735
         End
         Begin VB.TextBox Thrust_T2_Duration 
            Height          =   285
            Left            =   1200
            TabIndex        =   85
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox Thrust_T2_Strength 
            Height          =   285
            Left            =   1200
            TabIndex        =   84
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox Thrust_T2_Turnon_Iter 
            Height          =   285
            Left            =   1200
            TabIndex        =   83
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton Thrust_T2_Mode_OptionR 
            Height          =   255
            Left            =   1800
            TabIndex        =   82
            Top             =   2640
            Width           =   255
         End
         Begin VB.OptionButton Thrust_T2_Mode_OptionA 
            Height          =   255
            Left            =   1800
            TabIndex        =   81
            Top             =   2280
            Width           =   375
         End
         Begin VB.CheckBox Thrust_T2_Enable_Check 
            Caption         =   "Check1"
            Height          =   255
            Left            =   1200
            TabIndex        =   73
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label33 
            Caption         =   "Relative"
            Height          =   255
            Left            =   960
            TabIndex        =   80
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label Label32 
            Caption         =   "Absolute"
            Height          =   255
            Left            =   960
            TabIndex        =   79
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label Label31 
            Caption         =   "Mode"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label Label30 
            Caption         =   "Skew"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Label29 
            Caption         =   "Duration"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label28 
            Caption         =   "Strength"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label27 
            Caption         =   "Time"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label26 
            Caption         =   "Enable"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Thruster_T1_Frame 
         Caption         =   "T1"
         Height          =   3015
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   2295
         Begin VB.OptionButton Thrust_T1_Mode_OptionA 
            Height          =   375
            Left            =   1800
            TabIndex        =   71
            Top             =   2160
            Width           =   195
         End
         Begin VB.OptionButton Thrust_T1_Mode_OptionR 
            Height          =   195
            Left            =   1800
            TabIndex        =   70
            Top             =   2520
            Width           =   255
         End
         Begin VB.TextBox Thrust_T1_Strength 
            Height          =   285
            Left            =   1080
            TabIndex        =   66
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox Thrust_T1_Duration 
            Height          =   285
            Left            =   1080
            TabIndex        =   65
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox Thrust_T1_Skew 
            Height          =   285
            Left            =   1080
            TabIndex        =   64
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox Thrust_T1_Turnon_Iter 
            Height          =   285
            Left            =   1080
            TabIndex        =   63
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox Thrust_T1_Enable_Check 
            Caption         =   "Check1"
            Height          =   255
            Left            =   1080
            TabIndex        =   62
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label25 
            Caption         =   "Relative"
            Height          =   255
            Left            =   960
            TabIndex        =   69
            Top             =   2520
            Width           =   735
         End
         Begin VB.Label Label24 
            Caption         =   "Absolute"
            Height          =   255
            Left            =   960
            TabIndex        =   68
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label23 
            Caption         =   "Mode"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label Label22 
            Caption         =   "Strength"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label21 
            Caption         =   "Duration"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "Skew"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label19 
            Caption         =   "Time"
            Height          =   375
            Left            =   120
            TabIndex        =   58
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "Enable"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   855
         End
      End
   End
   Begin VB.CheckBox B_Check 
      Caption         =   "Check4"
      Height          =   195
      Left            =   960
      TabIndex        =   51
      Top             =   2880
      Width           =   255
   End
   Begin VB.CheckBox I_Check 
      Caption         =   "Check3"
      Height          =   255
      Left            =   960
      TabIndex        =   50
      Top             =   2280
      Width           =   255
   End
   Begin VB.CheckBox D_Check 
      Caption         =   "Check2"
      Height          =   255
      Left            =   960
      TabIndex        =   49
      Top             =   1680
      Width           =   255
   End
   Begin VB.CheckBox A_Check 
      Caption         =   "Check1"
      Height          =   255
      Left            =   960
      TabIndex        =   48
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox Object_Radius 
      Height          =   375
      Left            =   2280
      TabIndex        =   47
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox Seconds_per_Iter 
      Height          =   375
      Left            =   2280
      TabIndex        =   46
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset to Default"
      Height          =   495
      Left            =   4320
      TabIndex        =   42
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save Params"
      Height          =   495
      Left            =   120
      TabIndex        =   41
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Get Params"
      Height          =   495
      Left            =   2040
      TabIndex        =   40
      Top             =   7920
      Width           =   1935
   End
   Begin VB.TextBox B_delay 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   39
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Bmas 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   37
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox B_vely 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   36
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox B_velx 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   35
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox B_posy 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   34
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox B_posx 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   33
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox I_delay 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   29
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Imas 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   28
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox I_vely 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   27
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox I_velx 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   26
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox I_posy 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   25
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox I_posx 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   24
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Iter_Count 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   22
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Dmas 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   20
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Amas 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   19
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox D_vely 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   17
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox A_vely 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   16
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox D_velx 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   15
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox A_velx 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   14
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox G_const 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run Simulation"
      Height          =   495
      Left            =   6960
      TabIndex        =   9
      Top             =   7920
      Width           =   2055
   End
   Begin VB.TextBox D_posy 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox D_posx 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox A_posy 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox A_posx 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Param_Desc 
      Height          =   255
      Left            =   240
      TabIndex        =   53
      Top             =   7440
      Width           =   8895
   End
   Begin VB.Label Label17 
      Caption         =   "Object Radius"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   45
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "Seconds per Iteration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   44
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "Active"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   43
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "Delay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   38
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "Blip"
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   120
      TabIndex        =   32
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Initial Velocity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   31
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Initial Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   30
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Driftor2 
      Caption         =   "Intrudor"
      ForeColor       =   &H0080FF80&
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Total Iterations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   18
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Gravitational Constant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Vel y"
      Height          =   375
      Left            =   5640
      TabIndex        =   11
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Vel x"
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Driftor"
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Attractor"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Pos y"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Pos x"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form_Params"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call Reset_Params
End Sub

Private Sub Command1_Click()
    Form_Display.Show
    Form_Display.SetFocus
    Form_Display.Run_Simulation

End Sub

Private Sub Command2_Click()
    Call Reset_Params
End Sub


Private Sub Reset_Params()

    Form_Params.A_Check.Value = Checked
    Form_Params.A_posx = 100
    Form_Params.A_posy = 10
    Form_Params.A_velx = 0
    Form_Params.A_vely = 0.1
    Form_Params.Amas = 500
    
    Form_Params.D_Check.Value = Checked
    Form_Params.D_posx = -150
    Form_Params.D_posy = 50
    Form_Params.D_velx = 0
    Form_Params.D_vely = -4.25
    Form_Params.Dmas = 20.1
    
    Form_Params.I_Check.Value = Checked
    Form_Params.I_posx = -280
    Form_Params.I_posy = 280
    Form_Params.I_velx = -0.5
    Form_Params.I_vely = -1.65
    Form_Params.Imas = 12.1
    Form_Params.I_delay = 30000
    
    Form_Params.B_Check.Value = Checked
    Form_Params.B_posx = -290
    Form_Params.B_posy = 290
    Form_Params.B_velx = -0.5
    Form_Params.B_vely = -1.65
    Form_Params.Bmas = 0
    Form_Params.B_delay = 3000

    Form_Params.Iter_Count = 300000
    Form_Params.Seconds_per_Iter = 0.02
    Form_Params.Object_Radius = 0.001
    Form_Params.G_const = 12.5

End Sub

