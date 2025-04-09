VERSION 5.00
Begin VB.Form Form_Display 
   BackColor       =   &H80000008&
   Caption         =   "Form1"
   ClientHeight    =   9060
   ClientLeft      =   2445
   ClientTop       =   1935
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   11325
   Begin VB.CommandButton Command2 
      Caption         =   "Stop Simulation"
      Height          =   495
      Left            =   6240
      TabIndex        =   1
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Simulation"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   8400
      Width           =   1755
   End
End
Attribute VB_Name = "Form_Display"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim TimeInt As Variant

    Dim Gconst As Variant
    
    'Ranges
    Dim Range_ad As Variant
    Dim Range_ad_x As Variant
    Dim Range_ad_y As Variant
    
    Dim Range_id As Variant
    Dim Range_id_x As Variant
    Dim Range_id_y As Variant

    Dim Range_ia As Variant
    Dim Range_ia_x As Variant
    Dim Range_ia_y As Variant

    Dim Range_ab As Variant
    Dim Range_ab_x As Variant
    Dim Range_ab_y As Variant

    Dim Range_db As Variant
    Dim Range_db_x As Variant
    Dim Range_db_y As Variant

    Dim Range_ib As Variant
    Dim Range_ib_x As Variant
    Dim Range_ib_y As Variant

        
    ' Parameters for the Attractor
    Dim Apos_x As Variant
    Dim Apos_y As Variant
    Dim Avel_x As Variant
    Dim Avel_y As Variant
    Dim Amas As Variant
    
    ' Parameters for the Driftor
    Dim Dpos_x As Variant
    Dim Dpos_y As Variant
    Dim Dvel_x As Variant
    Dim Dvel_y As Variant
    Dim Dmas As Variant
   
    ' Parameters for the Intrudor
    Dim Ipos_x As Variant
    Dim Ipos_y As Variant
    Dim Ivel_x As Variant
    Dim Ivel_y As Variant
    Dim Imas As Variant
    Dim I_delay As Long

    
    ' Parameters for the Blip
    Dim Bpos_x As Variant
    Dim Bpos_y As Variant
    Dim Bvel_x As Variant
    Dim Bvel_y As Variant
    Dim Bmas As Variant
    Dim B_delay As Long

    

    Dim Stop_Flag As String
    
   'deltas of positions and velocities
    Dim del_Dposx_a As Variant
    Dim del_Dposy_a As Variant
    Dim del_Dvelx_a As Variant
    Dim del_Dvely_a As Variant
    Dim del_Dposx_i As Variant
    Dim del_Dposy_i As Variant
    Dim del_Dvelx_i As Variant
    Dim del_Dvely_i As Variant

    Dim del_Aposx_d As Variant
    Dim del_Aposy_d As Variant
    Dim del_Avelx_d As Variant
    Dim del_Avely_d As Variant
    Dim del_Aposx_i As Variant
    Dim del_Aposy_i As Variant
    Dim del_Avelx_i As Variant
    Dim del_Avely_i As Variant

    
    Dim del_Iposx_a As Variant
    Dim del_Iposy_a As Variant
    Dim del_Ivelx_a As Variant
    Dim del_Ively_a As Variant
    Dim del_Iposx_d As Variant
    Dim del_Iposy_d As Variant
    Dim del_Ivelx_d As Variant
    Dim del_Ively_d As Variant
 
    Dim del_Bposx_a As Variant
    Dim del_Bposy_a As Variant
    Dim del_Bvelx_a As Variant
    Dim del_Bvely_a As Variant
    Dim del_Bposx_d As Variant
    Dim del_Bposy_d As Variant
    Dim del_Bvelx_d As Variant
    Dim del_Bvely_d As Variant
    Dim del_Bposx_i As Variant
    Dim del_Bposy_i As Variant
    Dim del_Bvelx_i As Variant
    Dim del_Bvely_i As Variant

    
Private Sub Command1_Click()
    Form_Params.SetFocus
End Sub


Private Sub Command2_Click()
    Stop_Flag = "Y"
End Sub


Public Sub Run_Simulation()

    Dim T As Long
    
    'Constants
    Dim Iter_Count As Long
            
    Me.ScaleMode = 1 ' set to user scale
    
    ' to make 0,0 in the middle of the screen
    Me.ScaleLeft = -500      'Max left is -500
    Me.ScaleWidth = 1000     'Max right is 500
    Me.ScaleTop = 500        'Max top is 500
    Me.ScaleHeight = -1000    'Max bottom is -500
    
    
    'Position of the Attractor
    Apos_x = Form_Params.A_posx
    Apos_y = Form_Params.A_posy
    'Velocity of the Atractor (in pixels/sec)
    Avel_x = Form_Params.A_velx
    Avel_y = Form_Params.A_vely
    'Size of the Attractor
    Amas = Form_Params.Amas
    
    'Position of the Driftor
    Dpos_x = Form_Params.D_posx
    Dpos_y = Form_Params.D_posy
    'Velocity of the Driftor (in pixels/sec)
    Dvel_x = Form_Params.D_velx
    Dvel_y = Form_Params.D_vely
    'Size of the Driftor
    Dmas = Form_Params.Dmas
    
    'Position of the Intrudor
    Ipos_x = Form_Params.I_posx
    Ipos_y = Form_Params.I_posy
    'Velocity of the Intrudor (in pixels/sec)
    Ivel_x = Form_Params.I_velx
    Ivel_y = Form_Params.I_vely
    'Size of the Intrudor
    Imas = Form_Params.Imas

    'Position of the Blip
    Bpos_x = Form_Params.B_posx
    Bpos_y = Form_Params.B_posy
    'Velocity of the Blip (in pixels/sec)
    Bvel_x = Form_Params.B_velx
    Bvel_y = Form_Params.B_vely
    'Size of the Blip (usually zero mass)
    Bmas = Form_Params.Bmas

    
    'Set time interval per tick
    TimeInt = 0.008
    Iter_Count = Form_Params.Iter_Count
    'Set Constants
    Gconst = Form_Params.G_const
    ' Delay before the intrudor appears
    I_delay = Form_Params.I_delay
        ' Delay before the blip appears
    B_delay = Form_Params.B_delay

   
    ' Clear the screen
    Me.Cls
    
    'Draw a Red box 10 inside our form to prove we have the co-ordinates correct
    Me.Line (-490, -490)-(490, 490), RGB(255, 0, 0), B
        
    'Draw the origin
    Me.Line (-10, 0)-(10, 0), RGB(255, 255, 255) ' Draw Origin
    Me.Line (0, -10)-(0, 10), RGB(255, 255, 255) ' Draw Origin
                
    'Draw the attractor's initial position as a circle
    Me.Circle (Apos_x, Apos_y), 5, RGB(255, 255, 0)
    'Draw the driftor's initial position as a circle
    Me.Circle (Dpos_x, Dpos_y), 5, RGB(0, 100, 255)
    'Draw the intrudor's initial position as a circle
    Me.Circle (Ipos_x, Ipos_y), 5, RGB(100, 255, 0)
    'Draw the blip's initial position as a circle
    Me.Circle (Bpos_x, Bpos_y), 5, RGB(2550, 0, 255)


    'Initialize the stop flag
    Stop_Flag = "N"

    ' Calculate the curves and print them to the screen
    For T = 1 To Iter_Count
        ' Draw the Driftor's position on the screen
        Me.PSet (Dpos_x, Dpos_y), RGB(0, 100, 255) ' Draw Bluish Pixel
        ' Draw the attractor's position on the screen
        Me.PSet (Apos_x, Apos_y), RGB(255, 255, 0) ' Draws yellowish point
        ' Draw the intrudor's position on the screen
        Me.PSet (Ipos_x, Ipos_y), RGB(100, 255, 0) ' Draws greenish point
        
        ' Draw the blip's position on the screen
        If Form_Params.B_Check.Value = Checked Then
            Me.PSet (Bpos_x, Bpos_y), RGB(255, 0, 255) ' Draws mauvish point
        End If
  
        Call Calc_Ranges
        
        If Form_Params.A_Check.Value = Checked Then
            Call Motion_Due_to_Attractor
        End If
    
        If Form_Params.D_Check.Value = Checked Then
            Call Motion_Due_to_Driftor
        End If

        If T > I_delay Then
            If Form_Params.I_Check.Value = Checked Then
                Call Motion_Due_to_Intrudor
          End If
        End If
    
        Dvel_x = Dvel_x + del_Dvelx_i + del_Dvelx_a
        Dvel_y = Dvel_y + del_Dvely_i + del_Dvely_a
        Dpos_x = Dpos_x + del_Dposx_i + del_Dposx_a + (Dvel_x * TimeInt)
        Dpos_y = Dpos_y + del_Dposy_i + del_Dposy_a + (Dvel_y * TimeInt)
        
        
        Avel_x = Avel_x + del_Avelx_i + del_Avelx_d
        Avel_y = Avel_y + del_Avely_i + del_Avely_d
        Apos_x = Apos_x + del_Aposx_i + del_Aposx_d + (Avel_x * TimeInt)
        Apos_y = Apos_y + del_Aposy_i + del_Aposy_d + (Avel_y * TimeInt)

        If T > I_delay Then
            Ivel_x = Ivel_x + del_Ivelx_a + del_Ivelx_d
            Ivel_y = Ivel_y + del_Ively_a + del_Ively_d
            Ipos_x = Ipos_x + del_Iposx_a + del_Iposx_d + (Ivel_x * TimeInt)
            Ipos_y = Ipos_y + del_Iposy_a + del_Iposy_d + (Ivel_y * TimeInt)
        End If

        If T > B_delay Then
            Bvel_x = Bvel_x + del_Bvelx_a + del_Bvelx_d + del_Bvelx_i
            Bvel_y = Bvel_y + del_Bvely_a + del_Bvely_d + del_Bvely_i
            Bpos_x = Bpos_x + del_Bposx_a + del_Bposx_d + del_Bposx_i + (Bvel_x * TimeInt)
            Bpos_y = Bpos_y + del_Bposy_a + del_Bposy_d + del_Bposx_i + (Bvel_y * TimeInt)
        End If

           
        If Stop_Flag = "Y" Then
            'T = 10000
            Me.Circle (0, 0), 15, RGB(255, 0, 0)

        End If
        
    Next
    
    
End Sub


Private Sub Calc_Ranges()
        
        ' Calculate the ranges
        Range_ad_x = (Dpos_x - Apos_x)
        Range_ad_y = (Dpos_y - Apos_y)
        Range_ad = Sqr(Range_ad_x * Range_ad_x + Range_ad_y * Range_ad_y)
  
        Range_id_x = (Dpos_x - Ipos_x)
        Range_id_y = (Dpos_y - Ipos_y)
        Range_id = Sqr(Range_id_x * Range_id_x + Range_id_y * Range_id_y)
        
        Range_ia_x = (Apos_x - Ipos_x)
        Range_ia_y = (Apos_y - Ipos_y)
        Range_ia = Sqr(Range_ia_x * Range_ia_x + Range_ia_y * Range_ia_y)

        Range_ab_x = (Bpos_x - Apos_x)
        Range_ab_y = (Bpos_y - Apos_y)
        Range_ab = Sqr(Range_ab_x * Range_ab_x + Range_ab_y * Range_ab_y)

        Range_ib_x = (Bpos_x - Ipos_x)
        Range_ib_y = (Bpos_y - Ipos_y)
        Range_ib = Sqr(Range_ib_x * Range_ib_x + Range_ib_y * Range_ib_y)

        Range_db_x = (Bpos_x - Dpos_x)
        Range_db_y = (Bpos_y - Dpos_y)
        Range_db = Sqr(Range_db_x * Range_db_x + Range_db_y * Range_db_y)


End Sub

Private Sub Motion_Due_to_Attractor()
        
            'Variables
    
    Dim Accel As Variant
    Dim Accel_x As Variant
    Dim Accel_y As Variant
    Dim Bearing As Variant
        
        'Bearing = Asin(Range_y / Range)

        
        'Calculate the motion of the driftor due to the attractor
        Accel = Gconst * Amas / (Range_ad * Range_ad)
        
        Accel_x = -1 * (Range_ad_x / Range_ad) * Accel
        Accel_y = -1 * (Range_ad_y / Range_ad) * Accel
        
        del_Dvelx_a = (Accel_x * TimeInt)

        del_Dvely_a = (Accel_y * TimeInt)

        del_Dposx_a = (Accel_x * TimeInt * TimeInt) / 2
        
        del_Dposy_a = (Accel_y * TimeInt * TimeInt) / 2
        
                 
        'Calculate the motion of the intrudor due to the attractor
       
        Accel = Gconst * Amas / (Range_ia * Range_ia)
        
        Accel_x = 1 * (Range_ia_x / Range_ia) * Accel
        Accel_y = 1 * (Range_ia_y / Range_ia) * Accel
        
        del_Ivelx_a = (Accel_x * TimeInt)

        del_Ively_a = (Accel_y * TimeInt)

        del_Iposx_a = (Accel_x * TimeInt * TimeInt) / 2
        
        del_Iposy_a = (Accel_y * TimeInt * TimeInt) / 2
        
        'Calculate the motion of the blip due to the attractor
       
        Accel = Gconst * Amas / (Range_ab * Range_ab)
        
        Accel_x = -1 * (Range_ab_x / Range_ab) * Accel
        Accel_y = -1 * (Range_ab_y / Range_ab) * Accel
        
        del_Bvelx_a = (Accel_x * TimeInt)

        del_Bvely_a = (Accel_y * TimeInt)
        
        del_Bposx_a = (Accel_x * TimeInt * TimeInt) / 2
        
        del_Bposy_a = (Accel_y * TimeInt * TimeInt) / 2

        

End Sub

Private Sub Motion_Due_to_Driftor()

            'Variables
    
    Dim Accel As Variant
    Dim Accel_x As Variant
    Dim Accel_y As Variant
    Dim Bearing As Variant
        
        'Bearing = Asin(Range_y / Range)

       'Calculate the motion of the attractor due to the driftor
       
        Accel = Gconst * Dmas / (Range_ad * Range_ad)
        
        Accel_x = 1 * (Range_ad_x / Range_ad) * Accel
        Accel_y = 1 * (Range_ad_y / Range_ad) * Accel
        
        del_Avelx_d = (Accel_x * TimeInt)

        del_Avely_d = (Accel_y * TimeInt)
 
        del_Aposx_d = (Accel_x * TimeInt * TimeInt) / 2
        
        del_Aposy_d = (Accel_y * TimeInt * TimeInt) / 2
        

        'Calculate the motion of the intrudor due to the driftor
        
        Accel = Gconst * Dmas / (Range_id * Range_id)
        
        Accel_x = 1 * (Range_id_x / Range_id) * Accel
        Accel_y = 1 * (Range_id_y / Range_id) * Accel
        
        
        del_Ivelx_d = (Accel_x * TimeInt)

        del_Ively_d = (Accel_y * TimeInt)
        
        del_Iposx_d = (Accel_x * TimeInt * TimeInt) / 2
        
        del_Iposy_d = (Accel_y * TimeInt * TimeInt) / 2
        
        'Calculate the motion of the blip due to the driftor
        
        Accel = Gconst * Dmas / (Range_db * Range_db)
        
        Accel_x = -1 * (Range_db_x / Range_db) * Accel
        Accel_y = -1 * (Range_db_y / Range_db) * Accel
                
        del_Bvelx_d = (Accel_x * TimeInt)

        del_Bvely_d = (Accel_y * TimeInt)
        
        del_Bposx_d = (Accel_x * TimeInt * TimeInt) / 2
        
        del_Bposy_d = (Accel_y * TimeInt * TimeInt) / 2

End Sub


Private Sub Motion_Due_to_Intrudor()

    
    Dim Accel As Variant
    Dim Accel_x As Variant
    Dim Accel_y As Variant
    Dim Bearing As Variant

        
        'Calculate the motion of the driftor due to the intrudor
        Accel = Gconst * Imas / (Range_id * Range_id)
        
        Accel_x = -1 * (Range_id_x / Range_id) * Accel
        Accel_y = -1 * (Range_id_y / Range_id) * Accel
        
        del_Dvelx_i = (Accel_x * TimeInt)

        del_Dvely_i = (Accel_y * TimeInt)
        
        del_Dposx_i = (Accel_x * TimeInt * TimeInt) / 2
        
        del_Dposy_i = (Accel_y * TimeInt * TimeInt) / 2
        
                
        'Calculate the motion of the attractor due to the intrudor
       
        Accel = Gconst * Imas / (Range_ia * Range_ia)
        
        Accel_x = -1 * (Range_ia_x / Range_ia) * Accel
        Accel_y = -1 * (Range_ia_y / Range_ia) * Accel
        
        del_Avelx_i = (Accel_x * TimeInt)

        del_Avely_i = (Accel_y * TimeInt)
        
        del_Aposx_i = (Accel_x * TimeInt * TimeInt) / 2
        
        del_Aposy_i = (Accel_y * TimeInt * TimeInt) / 2
        
        'Calculate the motion of the blip due to the intrudor
       
        Accel = Gconst * Imas / (Range_ib * Range_ib)
        
        Accel_x = -1 * (Range_ib_x / Range_ib) * Accel
        Accel_y = -1 * (Range_ib_y / Range_ib) * Accel
        
        del_Bvelx_i = (Accel_x * TimeInt)

        del_Bvely_i = (Accel_y * TimeInt)
        
        del_Bposx_i = (Accel_x * TimeInt * TimeInt) / 2
        
        del_Bposy_i = (Accel_y * TimeInt * TimeInt) / 2
   

End Sub

