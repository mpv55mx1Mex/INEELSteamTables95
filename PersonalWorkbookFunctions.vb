Option Explicit

Dim StmTbl As SteamTables

Function TSubl(P As Double) As Double
    Set StmTbl = NewSteamTables
    TSubl = StmTbl.PSubl(P)
End Function

Function PSubl(T As Double) As Double
    Set StmTbl = NewSteamTables
    PSubl = StmTbl.PSubl(T)
End Function

Function TIce1(P As Double) As Double
    Set StmTbl = NewSteamTables
    TIce1 = StmTbl.TIce1(P)
End Function

Function PIce1(T As Double) As Double
    Set StmTbl = NewSteamTables
    PIce1 = StmTbl.PIce1(T)
End Function


Function TIce3(P As Double) As Double
    Set StmTbl = NewSteamTables
    TIce3 = StmTbl.TIce3(P)
End Function

Function PIce3(T As Double) As Double
    Set StmTbl = NewSteamTables
    PIce3 = StmTbl.PIce3(T)
End Function


Function TIce5(P As Double) As Double
    Set StmTbl = NewSteamTables
    TIce5 = StmTbl.TIce5(P)
End Function

Function PIce5(T As Double) As Double
    Set StmTbl = NewSteamTables
    PIce5 = StmTbl.PIce5(T)
End Function


Function TIce6(P As Double) As Double
    Set StmTbl = NewSteamTables
    TIce6 = StmTbl.TIce6(P)
End Function

Function PIce6(T As Double) As Double
    Set StmTbl = NewSteamTables
    PIce6 = StmTbl.PIce6(T)
End Function


Function TIce7(P As Double) As Double
    Set StmTbl = NewSteamTables
    TIce7 = StmTbl.TIce7(P)
End Function

Function PIce7(T As Double) As Double
    Set StmTbl = NewSteamTables
    PIce7 = StmTbl.PIce7(T)
End Function

Function TMelt(P As Double) As Double
    Set StmTbl = NewSteamTables
    TMelt = StmTbl.TMelt(P)
End Function

Function PMelt(T As Double) As Double
    Set StmTbl = NewSteamTables
    PMelt = StmTbl.PMelt(T)
End Function

Public Function TSat(Press As Double) As Double
    Set StmTbl = NewSteamTables
    Dim props As WtrProps
    
    StmTbl.SepBoundary = WtrSepBoundary.Saturation
    StmTbl.Pressure = Press
    StmTbl.PUpdate
    props = StmTbl.props
    
    TSat = props.temp
End Function

Public Function PSat(temp As Double) As Double
    Set StmTbl = NewSteamTables
    Dim props As WtrProps
    
    StmTbl.SepBoundary = WtrSepBoundary.Saturation
    StmTbl.Temperature = temp
    StmTbl.TUpdate
    props = StmTbl.props
    
    PSat = props.Press
End Function

Public Function T_Ext(Press As Double) As Double
    Set StmTbl = NewSteamTables
    Dim prop As WtrProps

    StmTbl.SepBoundary = WtrSepBoundary.CriticalIsochor
    StmTbl.Pressure = Press
    StmTbl.PUpdate
    prop = StmTbl.props
    
    T_Ext = prop.temp
End Function

Public Function P_Ext(temp As Double) As Double
    Set StmTbl = NewSteamTables
    Dim prop As WtrProps

    StmTbl.SepBoundary = WtrSepBoundary.CriticalIsochor
    StmTbl.Temperature = temp
    StmTbl.TUpdate
    prop = StmTbl.props
    
    P_Ext = prop.Press
End Function


Public Function TMinVol(Press As Double) As Double
    Set StmTbl = NewSteamTables
    Dim prop As WtrProps

    StmTbl.SepBoundary = WtrSepBoundary.MinimumVolume
    StmTbl.Pressure = Press
    StmTbl.PUpdate
    prop = StmTbl.props
    TMinVol = prop.temp
End Function

Public Function PMinVol(temp As Double) As Double
    Set StmTbl = NewSteamTables
    Dim prop As WtrProps

    StmTbl.SepBoundary = WtrSepBoundary.MinimumVolume
    StmTbl.Temperature = temp
    StmTbl.TUpdate
    prop = StmTbl.props
    PMinVol = prop.Press
End Function


Public Function SatPropsT(temp As Double) As Double()
    Set StmTbl = NewSteamTables

    Dim props As WtrProps

    StmTbl.SepBoundary = WtrSepBoundary.Saturation
    StmTbl.Temperature = temp
    StmTbl.TUpdate

    props = StmTbl.props

    SatPropsT = AssignProps(props)
End Function

Public Function SatPropsP(Press As Double) As Double()
    Set StmTbl = NewSteamTables

    Dim props As WtrProps

    StmTbl.SepBoundary = WtrSepBoundary.Saturation
    StmTbl.Pressure = Press
    StmTbl.PUpdate

    props = StmTbl.props

    SatPropsP = AssignProps(props)
End Function



Public Function NSatPropT(Nprop As Integer, Nphs As Integer, temp As Double) As Double

    Dim props() As Double
    
    props = SatPropsT(temp)
  
    If (Nphs = 1) Then
        NSatPropT = props(Nprop, 1)
    ElseIf (Nphs = 2) Then
        NSatPropT = props(Nprop, 2)
    Else
        NSatPropT = -1
    End If
End Function

Public Function NSatPropP(Nprop As Integer, Nphs As Integer, Press As Double) As Double

    Dim props() As Double
    
    props = SatPropsP(Press)
  
    If (Nphs = 1) Then
        NSatPropsT = props(Nprop, 1)
    ElseIf (Nphs = 2) Then
        NSatPropsT = props(Nprop, 2)
    Else
        NSatPropsT = -1
    End If
End Function


Public Function PropsTP(temp As Double, Press As Double) As Double()

    Set StmTbl = NewSteamTables
    Dim props As WtrProps

    StmTbl.Temperature = temp
    StmTbl.Pressure = Press
    StmTbl.TPUpdate

    props = StmTbl.props

    PropsTP = AssignProps(props)
End Function

Public Function NPropTP(Nprop As Integer, Nphs As Integer, temp As Double, Press As Double) As Double

    Dim props() As Double
    
    props = PropsTP(temp, Press)
  
    If (Nphs = 1) Then
        NPropTP = props(Nprop, 1)
    ElseIf (Nphs = 2) Then
        NPropTP = props(Nprop, 2)
    Else
        NPropTP = -1
    End If

End Function


Public Function AssignProps(props As WtrProps) As Double()
    'Assign the steam tables liquid and vapor properties
    '
    Dim prop(1 To 23, 1 To 2) As Double

    Dim propsLiq As PhaseProps
    Dim propsVap As PhaseProps

    propsLiq = props.Liquid
    propsVap = props.Vapor

    prop(1, 1) = propsLiq.Volume
    prop(2, 1) = propsLiq.density
    prop(3, 1) = propsLiq.Zo
    prop(4, 1) = propsLiq.U
    prop(5, 1) = propsLiq.H
    prop(6, 1) = propsLiq.G
    prop(7, 1) = propsLiq.A
    prop(8, 1) = propsLiq.s
    prop(9, 1) = propsLiq.Cp
    prop(10, 1) = propsLiq.Cv
    prop(11, 1) = propsLiq.CTE
    prop(12, 1) = propsLiq.Ziso
    prop(13, 1) = propsLiq.VelS
    prop(14, 1) = propsLiq.dPdT
    prop(15, 1) = propsLiq.dTdV
    prop(16, 1) = propsLiq.dVdP
    prop(17, 1) = propsLiq.JTC
    prop(18, 1) = propsLiq.IJTC
    prop(19, 1) = propsLiq.Vis
    prop(20, 1) = propsLiq.ThrmCond
    prop(21, 1) = propsLiq.SurfTen
    prop(22, 1) = propsLiq.PrdNum
    prop(23, 1) = propsLiq.DielCons

    prop(1, 2) = propsVap.Volume
    prop(2, 2) = propsVap.density
    prop(3, 2) = propsVap.Zo
    prop(4, 2) = propsVap.U
    prop(5, 2) = propsVap.H
    prop(6, 2) = propsVap.G
    prop(7, 2) = propsVap.A
    prop(8, 2) = propsVap.s
    prop(9, 2) = propsVap.Cp
    prop(10, 2) = propsVap.Cv
    prop(11, 2) = propsVap.CTE
    prop(12, 2) = propsVap.Ziso
    prop(13, 2) = propsVap.VelS
    prop(14, 2) = propsVap.dPdT
    prop(15, 2) = propsVap.dTdV
    prop(16, 2) = propsVap.dVdP
    prop(17, 2) = propsVap.JTC
    prop(18, 2) = propsVap.IJTC
    prop(19, 2) = propsVap.Vis
    prop(20, 2) = propsVap.ThrmCond
    prop(21, 2) = propsVap.SurfTen
    prop(22, 2) = propsVap.PrdNum
    prop(23, 2) = propsVap.DielCons

    AssignProps = prop
End Function

