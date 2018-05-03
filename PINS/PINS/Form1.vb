Imports System.Drawing.Printing
Imports System
Imports System.IO
Imports System.Math
Imports NCalc
Imports NCalc.Expression

Public Class Form1
    Dim Deci As Integer
    Dim Comp, UPC As String
    Dim DP, inpDP, ToothNo, ThetaR, convert, Rad, Pitch, inpToothNo, DoP, inpDoP, AlphaR, PHA, inpPHA, PHAr, PA, NPA, CPA, PAr, inpASW, ASW, BCD, inpPA, ArcTTh, TranTTh, inpArcTTh, PCD, inpPCD, MPD, inpMPD, Theta, Rb, Rt, Ri, Dbase, Pang, MPD1, Doe, Beta, BetaR, H, Hr, Alpha, Twoc, Eo, Vol As Double

    Private Sub Units_Click(sender As Object, e As EventArgs) Handles Units.Click
        If convert = 25.4 Then
            Units.Text = "Inches"
            convert = 1
        Else
            Units.Text = "MM"
            convert = 25.4
        End If
        ChangeDisplay()
    End Sub

    Sub ChangeDisplay()
        TxtArcTTh.Text = (inpArcTTh * convert)
        Txt2ASW.Text = (inpASW * convert)
        Txt3Sop.Text = (inpDoP * convert)
        Txt4Sop.Text = (inpDoP * convert)

        TxtMPD.Text = (inpMPD * convert)
        Txt2MPD.Text = (inpMPD * convert)
        Txt3MPD.Text = (inpMPD * convert)
        Txt4MPD.Text = (inpMPD * convert)
    End Sub

    'Printing
    Private Sub PrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem.Click
        PrintDocument1.Print()

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        convert = 1
        Me.PrintPreviewDialog1 = New PrintPreviewDialog

        'Set the size, location, and name.
        Me.PrintPreviewDialog1.ClientSize = New System.Drawing.Size(400, 300)
        Me.PrintPreviewDialog1.Location = New System.Drawing.Point(29, 29)
        Me.PrintPreviewDialog1.Name = "PrintPreviewDialog1"

        ' Set the minimum size the dialog can be resized to.
        Me.PrintPreviewDialog1.MinimumSize = New System.Drawing.Size(375, 250)

        ' Set the UseAntiAlias property to true, which will allow the 
        ' operating system to smooth fonts.
        Me.PrintPreviewDialog1.UseAntiAlias = True
    End Sub

    Private Sub PrinterSetupToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrinterSetupToolStripMenuItem.Click
        Try
            PrintDialog1.Document = PrintDocument1
            PrintDialog1.PrinterSettings = PrintDocument1.PrinterSettings
            PrintDialog1.AllowSomePages = True
            If PrintDialog1.ShowDialog = DialogResult.OK Then
                PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings
                PrintDocument1.Print()
            End If
        Catch i As Exception
            MsgBox("An error has occured:" & i.ToString)
        End Try
    End Sub
    Private Sub PrintPreviewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintPreviewToolStripMenuItem.Click
        PrintPreviewDialog1.Document = PrintDocument1

        PrintPreviewDialog1.ShowDialog()
    End Sub


    'Helical will determine the normal and transverse arcthick and spadewidths

    'IF DP is being used then it will also mean NDP, CDP, NMOD, AND CMOD would need inputting through a dropdown box if Helical is selected. 

    Private Sub CheckHelical_CheckedChanged(sender As Object, e As EventArgs) Handles CheckHelical.CheckedChanged
        If CheckHelical.Checked = True Then
            CheckSpur.Checked = False
        End If

        ComboPitch.Items.Clear()
        ComboPitch.Text = ""
        ComboPitch.Items.Add("NDP")
        ComboPitch.Items.Add("CDP")
        ComboPitch.Items.Add("NMOD")
        ComboPitch.Items.Add("CMOD")

        Combo2Pitch.Items.Clear()
        Combo2Pitch.Text = ""
        Combo2Pitch.Items.Add("NDP")
        Combo2Pitch.Items.Add("CDP")
        Combo2Pitch.Items.Add("NMOD")
        Combo2Pitch.Items.Add("CMOD")

        ComboFn3Pitch.Items.Clear()
        ComboFn3Pitch.Text = ""
        ComboFn3Pitch.Items.Add("NDP")
        ComboFn3Pitch.Items.Add("CDP")
        ComboFn3Pitch.Items.Add("NMOD")
        ComboFn3Pitch.Items.Add("CMOD")

        Combo4Pitch.Items.Clear()
        Combo4Pitch.Text = ""
        Combo4Pitch.Items.Add("NDP")
        Combo4Pitch.Items.Add("CDP")
        Combo4Pitch.Items.Add("NMOD")
        Combo4Pitch.Items.Add("CMOD")

    End Sub

    Private Sub CheckSpur_Clicked(sender As Object, e As EventArgs) Handles CheckSpur.Click
        If CheckSpur.Checked = True Then
            CheckHelical.Checked = False
        End If

        ComboPitch.Items.Clear()
        ComboPitch.Text = ""
        ComboPitch.Items.Add("Diametral Pitch")
        ComboPitch.Items.Add("Module")


        Combo2Pitch.Items.Clear()
        Combo2Pitch.Text = ""
        Combo2Pitch.Items.Add("Diametral Pitch")
        Combo2Pitch.Items.Add("Module")

        ComboFn3Pitch.Items.Clear()
        ComboFn3Pitch.Text = ""
        ComboFn3Pitch.Items.Add("Diametral Pitch")
        ComboFn3Pitch.Items.Add("Module")

        Combo4Pitch.Items.Clear()
        Combo4Pitch.Text = ""
        Combo4Pitch.Items.Add("Diametral Pitch")
        Combo4Pitch.Items.Add("Module")

    End Sub

    'Private Sub CheckInternal_Clicked(sender As Object, e As EventArgs)
    '    If CheckInternal.Checked = True Then
    '        CheckExternal.Checked = False
    '    End If
    '    Fn3LblSop.Text = "Dimension under Pins"
    '    Fn4LblSop.Text = "Dimension under Pins"
    'End Sub

    'Private Sub CheckExternal_Clicked(sender As Object, e As EventArgs)
    '    If CheckExternal.Checked = True Then
    '        CheckInternal.Checked = False
    '    End If
    '    Fn3LblSop.Text = "Dimension over Pins"
    '    Fn4LblSop.Text = "Dimension over Pins"
    'End Sub

    Private Sub RadioFunc4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioFunc4.CheckedChanged
        PanelFunc3.Visible = True
        PanelFunc3.Location = New Point(12, 41)
        PanelFunc4.Visible = False
        PanelFunc2.Visible = False
        PanelFunc1.Visible = False
        Doe = 0
        Rt = 0
        Theta = 0
        ToothNo = 0
        PHA = 0
        ArcTTh = 0
        PCD = 0
        PA = 0
        MPD = 0
        inpToothNo = 0
        inpPHA = 0
        inpArcTTh = 0
        inpPCD = 0
        inpPA = 0
        inpMPD = 0
        DP = 0
        Alpha = 0
        Vol = 0
        AlphaR = 0
        Pitch = 0
        TxtPitch.Text = ""
        Txt2Pitch.Text = ""
        Txt3Pitch.Text = ""
        Txt4Pitch.Text = ""
        TxtDoe.Text = ""
        TxtRt.Text = ""
        TxtTheta.Text = ""
        TxtToothNo.Text = ""
        Txt2ToothNo.Text = ""
        TxtPHA.Text = ""
        Txt2PHA.Text = ""
        Txt3PHA.Text = ""
        Txt4PHA.Text = ""
        TxtArcTTh.Text = ""
        'TxtPCD.Text = ""
        'Txt2PCD.Text = ""
        'Txt3PCD.Text = ""
        'Txt4PCD.Text = ""
        TxtPA.Text = ""
        Txt3PA.Text = ""
        Txt4PA.Text = ""
        TxtMPD.Text = ""
        Txt2MPD.Text = ""
        Txt3MPD.Text = ""
        Txt4MPD.Text = ""
        TxtDoe.Text = ""
        TxtRt.Text = ""
        TxtTheta.Text = ""
    End Sub

    Private Sub RadioFunc3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioFunc3.CheckedChanged
        PanelFunc3.Visible = False
        PanelFunc4.Visible = True
        PanelFunc4.Location = New Point(12, 41)
        PanelFunc2.Visible = False
        PanelFunc1.Visible = False
        Doe = 0
        Rt = 0
        Theta = 0
        ToothNo = 0
        PHA = 0
        ArcTTh = 0
        PCD = 0
        PA = 0
        MPD = 0
        inpToothNo = 0
        inpPHA = 0
        inpArcTTh = 0
        inpPCD = 0
        inpPA = 0
        inpMPD = 0
        ASW = 0
        DP = 0
        Pitch = 0
        Alpha = 0
        Vol = 0
        AlphaR = 0
        TxtPitch.Text = ""
        Txt2Pitch.Text = ""
        Txt3Pitch.Text = ""
        Txt4Pitch.Text = ""
        TxtDoe.Text = ""
        TxtRt.Text = ""
        TxtTheta.Text = ""
        TxtToothNo.Text = ""
        Txt2ToothNo.Text = ""
        TxtPHA.Text = ""
        Txt2PHA.Text = ""
        Txt3PHA.Text = ""
        Txt4PHA.Text = ""
        TxtArcTTh.Text = ""
        'TxtPCD.Text = ""
        'Txt2PCD.Text = ""
        'Txt3PCD.Text = ""
        'Txt4PCD.Text = ""
        TxtPA.Text = ""
        Txt3PA.Text = ""
        Txt4PA.Text = ""
        TxtMPD.Text = ""
        Txt2MPD.Text = ""
        Txt3MPD.Text = ""
        Txt4MPD.Text = ""




    End Sub

    Private Sub RadioFunc2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioFunc2.CheckedChanged
        PanelFunc3.Visible = False
        PanelFunc4.Visible = False
        PanelFunc2.Visible = True
        PanelFunc2.Location = New Point(12, 41)
        PanelFunc1.Visible = False
        TxtPA.Visible = False
        ComboPA.Visible = False

        Doe = 0
        Rt = 0
        Theta = 0
        ToothNo = 0
        PHA = 0
        ArcTTh = 0
        PCD = 0
        ASW = 0
        PA = 0
        MPD = 0
        DP = 0
        Pitch = 0
        Alpha = 0
        Vol = 0
        AlphaR = 0
        TxtPitch.Text = ""
        Txt2Pitch.Text = ""
        Txt3Pitch.Text = ""
        Txt4Pitch.Text = ""
        inpToothNo = 0
        inpPHA = 0
        inpArcTTh = 0
        inpPCD = 0
        inpPA = 0
        inpMPD = 0
        TxtDoe.Text = ""
        TxtRt.Text = ""
        TxtTheta.Text = ""
        TxtToothNo.Text = ""
        Txt2ToothNo.Text = ""
        TxtPHA.Text = ""
        Txt2PHA.Text = ""
        Txt3PHA.Text = ""
        Txt4PHA.Text = ""
        TxtArcTTh.Text = ""
        'TxtPCD.Text = ""
        'Txt2PCD.Text = ""
        'Txt3PCD.Text = ""
        'Txt4PCD.Text = ""
        TxtPA.Text = ""
        Txt3PA.Text = ""
        Txt4PA.Text = ""
        TxtMPD.Text = ""
        Txt2MPD.Text = ""
        Txt3MPD.Text = ""
        Txt4MPD.Text = ""
        TxtDoe.Text = ""
        TxtRt.Text = ""
        TxtTheta.Text = ""
    End Sub

    Private Sub RadioFunc1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioFunc1.CheckedChanged
        PanelFunc3.Visible = False
        PanelFunc4.Visible = False
        PanelFunc2.Visible = False
        PanelFunc1.Visible = True

        TxtPA.Visible = True
        ComboPA.Visible = True


        Doe = 0
        Rt = 0
        Theta = 0
        ToothNo = 0
        PHA = 0
        ArcTTh = 0
        PCD = 0
        PA = 0
        MPD = 0
        ASW = 0
        DP = 0
        Pitch = 0
        TxtPitch.Text = ""
        Txt2Pitch.Text = ""
        Txt3Pitch.Text = ""
        Txt4Pitch.Text = ""
        inpToothNo = 0
        inpPHA = 0
        inpArcTTh = 0
        inpPCD = 0
        inpPA = 0
        inpMPD = 0
        TxtDoe.Text = ""
        TxtRt.Text = ""
        TxtTheta.Text = ""
        TxtToothNo.Text = ""
        Txt2ToothNo.Text = ""
        TxtPHA.Text = ""
        Txt2PHA.Text = ""
        Txt3PHA.Text = ""
        Txt4PHA.Text = ""
        TxtArcTTh.Text = ""
        'TxtPCD.Text = ""
        'Txt2PCD.Text = ""
        'Txt3PCD.Text = ""
        'Txt4PCD.Text = ""
        TxtPA.Text = ""
        Txt3PA.Text = ""
        Txt4PA.Text = ""
        TxtMPD.Text = ""
        Txt2MPD.Text = ""
        Txt3MPD.Text = ""
        Txt4MPD.Text = ""
        TxtDoe.Text = ""
        TxtRt.Text = ""
        TxtTheta.Text = ""
        Txt3Sop.Text = ""
        Txt4Sop.Text = ""
        Txt2ASW.Text = ""
    End Sub

    Private Sub DecimalPlacesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DecimalPlacesToolStripMenuItem.Click
        Deci = Val(InputBox("How many decimal places would you like to display?", "Decimal Places", "5"))
    End Sub

    Private Sub BtnCalc_Click(sender As Object, e As EventArgs) Handles BtnCalc.Click
        If RadioFunc1.Checked = True Then
            Opft()
        ElseIf RadioFunc2.Checked = True Then

            Upfs()
            'Do 2nd func
        ElseIf RadioFunc3.Checked = True Then


            'Do 3rd func 
        ElseIf RadioFunc4.Checked = True Then
          

            'Do 4th func
        Else
            MsgBox("Please select a function")
        End If
    End Sub
    Sub Upfs()
        Rad = 180 / PI
        PAr = PA / Rad
        H = 0
        MPD1 = MPD
        BetaR = Beta / Rad
        If Beta > 0.000001 Then

            H = Atan(Tan(Beta) * Cos(PA))

            MPD = inpMPD / Cos(H)
        End If

        BCD = PCD * Cos(PAr)
        Vol = ASW / PCD + (Tan(PA) - PA) - MPD / BCD
        Alpha = Ainv(Vol)
        Twoc = BCD / Cos(Alpha)

        Eo = ToothNo Mod 2

        If Eo > 0.1 Then
            Doe = Twoc * Cos(90 / ToothNo / Rad) + MPD1
        Else
            Doe = Twoc - MPD1
        End If
        Theta = Atan(Tan(Alpha) + MPD1 * Cos(H) / BCD)
        Rt = BCD / (2 * Cos(Theta))

        TxtTheta.Text = ThetaR
        TxtRt.Text = Rt
        TxtDoe.Text = Doe

    End Sub

    Sub Opft()
        Rad = 180 / PI
        MPD1 = inpMPD

        PAr = PA / Rad
        PHAr = PHA / 180 * PI
        'SHOULD THAT BE PHAr or InpPHA?
        Beta = PHAr
        BetaR = Beta / Rad
        PCD = ToothNo / Pitch
        'MsgBox(PCD)


        If ComboPA.SelectedItem = "NPA" Then
            PA = inpPA
        ElseIf ComboPA.SelectedItem = "CPA" Then
            PA = Atan(Tan(inpPA) / Cos(BetaR))
            MsgBox(PA)
        End If

        If ComboThick.SelectedItem = "Transverse Arc Thickness" Then
            ' TranTTh = PA = Atan(Tan(PA) / Cos(Beta))
            TranTTh = inpArcTTh
            'ArcTTh = inpArcTTh
        Else
            TranTTh = inpArcTTh / Cos(Beta)
        End If

        If Beta > 0.000001 Then

            H = Atan(Tan(Beta) * Cos(PA))

            MPD = inpMPD / Cos(H)
        End If
        '  MsgBox(PCD)
        BCD = PCD * Cos(PAr)

        Vol = TranTTh / PCD + MPD / BCD + (Tan(PAr) - PAr) - PI / ToothNo
        MsgBox("Vol = " & Vol)
        Alpha = Ainv(Vol)
        Twoc = BCD / Cos(Alpha)
        Eo = ToothNo Mod 2

        If Eo > 0.1 Then
            Doe = Twoc * Cos(90 / ToothNo / Rad) + MPD1
        Else
            Doe = Twoc + MPD1
        End If

        Theta = Atan(Tan(Alpha) - MPD1 * Cos(H) / BCD)
        Rt = BCD / (2 * Cos(Theta))
        ThetaR = Theta * 180 / PI

        TxtTheta.Text = ThetaR
        TxtRt.Text = Rt
        TxtDoe.Text = Doe

    End Sub

    Sub Input(ByRef Input As String)
        If Input = "" Then Exit Sub
        If Input.Contains("pi") Then Input = Input.Replace("pi", " 3.14159265358979")
        If Input.Contains("Pi") Then Input = Input.Replace("Pi", " 3.14159265358979")
        If Input.Contains("PI") Then Input = Input.Replace("PI", " 3.14159265358979")
        If Input.Contains("pI") Then Input = Input.Replace("pI", " 3.14159265358979")
        Dim exp As NCalc.Expression = New NCalc.Expression(Input, EvaluateOptions.IgnoreCase)
        Try
            Input = exp.Evaluate
        Catch ex As Exception
            Input = ""
        End Try
    End Sub

    'Private Sub CheckExternal_CheckedChanged(sender As Object, e As EventArgs)
    '    If CheckExternal.Checked = True Then
    '        CheckInternal.Checked = False
    '    End If
    'End Sub

    'Private Sub CheckInternal_CheckedChanged(sender As Object, e As EventArgs)
    '    If CheckInternal.Checked = True Then
    '        CheckExternal.Checked = False
    '    End If
    'End Sub

    'Private Sub CheckSpur_CheckedChanged(sender As Object, e As EventArgs)
    '    If CheckSpur.Checked = True Then
    '        CheckHelical.Checked = False
    '    End If
    'End Sub

    'Private Sub CheckHelical_CheckedChanged(sender As Object, e As EventArgs)
    '    If CheckHelical.Checked = True Then
    '        CheckSpur.Checked = False
    '    End If
    'End Sub

    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'FUNCTION 1 INPUTS
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Sub TxtToothNo_LostFocus(sender As Object, e As EventArgs) Handles TxtToothNo.LostFocus
        Input(TxtToothNo.Text)
        ToothNo = Val(TxtToothNo.Text)
        inpToothNo = Val(TxtToothNo.Text)
    End Sub

    Private Sub TxtPHA_LostFocus(sender As Object, e As EventArgs) Handles TxtPHA.LostFocus
        Input(TxtPHA.Text)
        PHA = Val(TxtPHA.Text)

        inpPHA = Val(TxtPHA.Text)
    End Sub

    Private Sub TxtArcTTh_LostFocus(sender As Object, e As EventArgs) Handles TxtArcTTh.LostFocus
        Input(TxtArcTTh.Text)
        ArcTTh = Val(TxtArcTTh.Text)

        inpArcTTh = Val(TxtArcTTh.Text) / convert
    End Sub

    'Private Sub TxtPCD_LostFocus(sender As Object, e As EventArgs)
    '    Input(TxtPCD.Text)
    '    PCD = Val(TxtPCD.Text)
    '    inpPCD = Val(TxtPCD.Text)/ convert
    'End Sub
    Private Sub TxtPA_LostFocus(sender As Object, e As EventArgs) Handles TxtPA.LostFocus
        Input(TxtPA.Text)
        PA = Val(TxtPA.Text)
        inpPA = Val(TxtPA.Text)
    End Sub

    Private Sub TxtMPD_LostFocus(sender As Object, e As EventArgs) Handles TxtMPD.LostFocus
        Input(TxtMPD.Text)
        MPD = Val(TxtMPD.Text)
        inpMPD = Val(TxtMPD.Text) / convert
    End Sub
    Private Sub TxtPitch_LostFocus(sender As Object, e As EventArgs) Handles TxtPitch.LostFocus
        Input(TxtPitch.Text)
        ' DP = Val(TxtPitch.Text)
        inpDP = Val(TxtPitch.Text)
        Dim CDP As Double
        If ComboPitch.SelectedItem = "Diametral Pitch" Then
            inpDP = DP

        ElseIf ComboPitch.SelectedItem = "Module" Then
            DP = 25.4 / inpDP
            MsgBox(DP)
            Pitch = DP
        ElseIf ComboPitch.SelectedItem = "NMOD" Then
            DP = 25.4 / inpDP
            Pitch = DP
        ElseIf ComboPitch.SelectedItem = "CDP" Then
            DP = inpDP / Cos(PHA)
            Pitch = DP
        ElseIf ComboPitch.SelectedItem = "CMOD" Then
            CDP = 25.4 / inpDP
            DP = CDP / Cos(PHA)
            Pitch = DP
        End If
    End Sub
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'FUNCTION 2 INPUTS
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub Txt2Pitch_LostFocus(sender As Object, e As EventArgs) Handles TxtPitch.LostFocus
        Input(TxtPitch.Text)
        DP = Val(TxtPitch.Text)
        inpDP = Val(TxtPitch.Text)
        Dim CDP As Double
        If Combo2Pitch.SelectedItem = "Diametral Pitch" Then
            inpDP = DP
            Pitch = DP
        ElseIf Combo2Pitch.SelectedItem = "Module" Then
            DP = 25.4 / inpDP
            Pitch = DP
        ElseIf Combo2Pitch.SelectedItem = "NMOD" Then
            DP = 25.4 / inpDP
            Pitch = DP
        ElseIf Combo2Pitch.SelectedItem = "CDP" Then
            DP = inpDP / Cos(PHA)
            Pitch = DP
        ElseIf Combo2Pitch.SelectedItem = "CMOD" Then
            CDP = 25.4 / inpDP
            DP = CDP / Cos(PHA)
            Pitch = DP
        End If
    End Sub

    Private Sub Txt2MPD_LostFocus(sender As Object, e As EventArgs) Handles Txt2MPD.LostFocus
        Input(Txt2MPD.Text)
        MPD = Val(Txt2MPD.Text)

        inpMPD = Val(Txt2MPD.Text) / convert
    End Sub

    'Private Sub Txt2PCD_LostFocus(sender As Object, e As EventArgs)
    '    Input(Txt2PCD.Text)
    '    PCD = Val(Txt2PCD.Text)

    '    inpPCD = Val(Txt2PCD.Text)
    'End Sub

    Private Sub Txt2ASW_LostFocus(sender As Object, e As EventArgs) Handles Txt2ASW.LostFocus
        Input(Txt2ASW.Text)
        ASW = Val(Txt2ASW.Text)

        inpASW = Val(Txt2ASW.Text) / convert
    End Sub

    Private Sub Txt2PHA_LostFocus(sender As Object, e As EventArgs) Handles Txt2PHA.LostFocus
        Input(Txt2PHA.Text)
        PHA = Val(Txt2PHA.Text)

        inpPHA = Val(Txt2PHA.Text)
    End Sub

    Private Sub Txt2ToothNo_LostFocus(sender As Object, e As EventArgs) Handles Txt2ToothNo.LostFocus
        Input(Txt2ToothNo.Text)
        ToothNo = Val(Txt2ToothNo.Text)

        inpToothNo = Val(Txt2ToothNo.Text)
    End Sub
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'FUNCTION 3 INPUTS
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub Txt3Pitch_LostFocus(sender As Object, e As EventArgs) Handles TxtPitch.LostFocus
        Input(TxtPitch.Text)
        DP = Val(TxtPitch.Text)
        inpDP = Val(TxtPitch.Text)
        Dim CDP As Double
        If ComboFn3Pitch.SelectedItem = "Diametral Pitch" Then
            inpDP = DP
            Pitch = DP
        ElseIf ComboFn3Pitch.SelectedItem = "Module" Then
            DP = 25.4 / inpDP
            Pitch = DP
        ElseIf ComboFn3Pitch.SelectedItem = "NMOD" Then
            DP = 25.4 / inpDP
            Pitch = DP
        ElseIf ComboFn3Pitch.SelectedItem = "CDP" Then
            DP = inpDP / Cos(PHA)
            Pitch = DP
        ElseIf ComboFn3Pitch.SelectedItem = "CMOD" Then
            CDP = 25.4 / inpDP
            DP = CDP / Cos(PHA)
            Pitch = DP
        End If
    End Sub

    Private Sub Txt3MPD_LostFocus(sender As Object, e As EventArgs) Handles Txt3MPD.LostFocus
        Input(Txt3MPD.Text)
        MPD = Val(Txt3MPD.Text)

        inpMPD = Val(Txt3MPD.Text) / convert
    End Sub

    'Private Sub Txt3PCD_LostFocus(sender As Object, e As EventArgs)
    '    Input(Txt3PCD.Text)
    '    PCD = Val(Txt3PCD.Text)

    '    inpPCD = Val(Txt3PCD.Text)
    'End Sub

    Private Sub Txt3Sop_LostFocus(sender As Object, e As EventArgs) Handles Txt3Sop.LostFocus
        Input(Txt3Sop.Text)
        DoP = Val(Txt3Sop.Text)

        inpDoP = Val(Txt3Sop.Text) / convert
    End Sub

    Private Sub Txt3PHA_LostFocus(sender As Object, e As EventArgs) Handles Txt3PHA.LostFocus
        Input(Txt3PHA.Text)
        PHA = Val(Txt3PHA.Text)

        inpPHA = Val(Txt3PHA.Text)
    End Sub

    Private Sub Txt3PA_LostFocus(sender As Object, e As EventArgs) Handles Txt3PA.LostFocus
        Input(Txt3PA.Text)
        PA = Val(Txt3PA.Text)

        inpPA = Val(Txt3PA.Text)
    End Sub

    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'FUNCTION 4 INPUTS
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub Txt4Pitch_LostFocus(sender As Object, e As EventArgs) Handles TxtPitch.LostFocus
        Input(TxtPitch.Text)
        DP = Val(TxtPitch.Text)
        inpDP = Val(TxtPitch.Text)

        Dim CDP As Double
        If Combo4Pitch.SelectedItem = "Diametral Pitch" Then
            inpDP = DP
            Pitch = DP
        ElseIf Combo4Pitch.SelectedItem = "Module" Then
            DP = 25.4 / inpDP
            Pitch = DP
        ElseIf Combo4Pitch.SelectedItem = "NMOD" Then
            DP = 25.4 / inpDP
            Pitch = DP
        ElseIf Combo4Pitch.SelectedItem = "CDP" Then
            DP = inpDP / Cos(PHA)
            Pitch = DP
        ElseIf Combo4Pitch.SelectedItem = "CMOD" Then
            CDP = 25.4 / inpDP
            DP = CDP / Cos(PHA)
            Pitch = DP
        End If
    End Sub

    Private Sub Txt4MPD_LostFocus(sender As Object, e As EventArgs) Handles Txt4MPD.LostFocus
        Input(Txt4MPD.Text)
        MPD = Val(Txt4MPD.Text)

        inpMPD = Val(Txt4MPD.Text) / convert
    End Sub

    'Private Sub Txt4PCD_LostFocus(sender As Object, e As EventArgs)
    '    Input(Txt4PCD.Text)
    '    PCD = Val(Txt4PCD.Text)

    '    inpPCD = Val(Txt4PCD.Text)
    'End Sub

    Private Sub Txt4Sop_LostFocus(sender As Object, e As EventArgs) Handles Txt4Sop.LostFocus
        Input(Txt4Sop.Text)
        DoP = Val(Txt4Sop.Text)

        inpDoP = Val(Txt4Sop.Text) / convert
    End Sub

    Private Sub Txt4PHA_LostFocus(sender As Object, e As EventArgs) Handles Txt4PHA.LostFocus
        Input(Txt4PHA.Text)
        PHA = Val(Txt4PHA.Text)

        inpPHA = Val(Txt4PHA.Text)
    End Sub

    Private Sub Txt4PA_LostFocus(sender As Object, e As EventArgs) Handles Txt4PA.LostFocus
        Input(Txt4PA.Text)
        PA = Val(Txt4PA.Text)

        inpPA = Val(Txt4PA.Text)
    End Sub
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Shared Function Inv(ByVal Ang As Double) As Double
        Inv = Tan(Ang) - Ang
    End Function

    Public Shared Function Ainv(ByVal Inv As Double) As Double
        Dim V1, V2, V3, V4 As Double
        V1 = Exp(Log(3 + Inv) / 3)
        V2 = Min(V1, 4 * Atan(1) / 2.000001)
        Do
            V3 = V2 - Tan(V2) + Inv
            V4 = V3 / Tan(V2) ^ 2 + V2
            If V2 - V4 < 0.000001 Then Exit Do
            V2 = V4
        Loop
        Ainv = V2
    End Function
End Class
