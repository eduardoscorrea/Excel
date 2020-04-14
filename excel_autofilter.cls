Sub teste()

ActiveWorkbook.Sheets("Projetos Concluídos").Activate
Range("C:AB").Select
Selection.AutoFilter
Dim crproduct(39) As String
Dim j As Integer
j = 0

crproduct(1) = "1138735"
crproduct(2) = "1140623"
crproduct(3) = "1141093"
crproduct(4) = "1138551"
crproduct(5) = "1138745"
crproduct(6) = "1138757"
crproduct(7) = "1141028"
crproduct(8) = "1141713"
crproduct(9) = "1142835"
crproduct(10) = "1141559"
crproduct(11) = "1139033"
crproduct(12) = "1113782"
crproduct(13) = "1142183"
crproduct(14) = "906558"
crproduct(15) = "1123902"
crproduct(16) = "1144281"
crproduct(17) = "1143815"
crproduct(18) = "1123854"
crproduct(19) = "1136041"
crproduct(20) = "1052569"
crproduct(21) = "1144783"
crproduct(22) = "1144433"
crproduct(23) = "1143667"
crproduct(24) = "1143705"
crproduct(25) = "1144835"
crproduct(26) = "1145957"
crproduct(27) = "1146445"
crproduct(28) = "1148277"
crproduct(29) = "1139095"
crproduct(30) = "868470"
crproduct(31) = "1143499"
crproduct(32) = "1149755"
crproduct(33) = "1149761"
crproduct(34) = "1149607"
crproduct(35) = "824471"
crproduct(36) = "872381"
crproduct(37) = "808130"
crproduct(38) = "1150533"
crproduct(39) = "1150489"

Application.ScreenUpdating = False

Selection.AutoFilter Field:=1, Criteria1:=Array(crproduct(0), crproduct(1), crproduct(2), crproduct(3), crproduct(4), crproduct(5), crproduct(6), crproduct(7), crproduct(8), crproduct(9), _
                                                 crproduct(10), crproduct(11), crproduct(12), crproduct(13), crproduct(14), crproduct(15), crproduct(16), crproduct(17), crproduct(18), crproduct(19), _
                                                 crproduct(20), crproduct(21), crproduct(22), crproduct(23), crproduct(24), crproduct(25), crproduct(26), crproduct(27), crproduct(28), crproduct(29), _
                                                 crproduct(30), crproduct(31), crproduct(32), crproduct(33), crproduct(34), crproduct(35), crproduct(36), crproduct(37), crproduct(38), crproduct(39)), _
                                                 Operator:=xlFilterValues
Application.ScreenUpdating = True

End Sub
                