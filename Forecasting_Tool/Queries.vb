Module Queries

    Public Function Hours_Per_Day(Source As String) As String
        Return "(" &
"SELECT " &
"ToolLink.WhatDay, ToolLink.FinanceID, ([SSS]*[P_SSS_Mins])/60 As SSS_P_Hours, ([SSS]*[N_SSS_Mins])/60 As SSS_N_Hours, ([SSS]*[C_SSS_Mins])/60 As SSS_C_Hours, " &
"([PSP]*[P_PSP_Mins])/60 As PSP_P_Hours, ([PSP]*[N_PSP_Mins])/60 As PSP_N_Hours, ([PSP]*[C_PSP_Mins])/60 As PSP_C_Hours, ([PSP2]*[P_PSP2_Mins])/60 As PSP2_P_Hours, " &
"([PSP2]*[N_PSP2_Mins])/60 As PSP2_N_Hours, ([PSP2]*[C_PSP2_Mins])/60 As PSP2_C_Hours, ([FUP]*[P_FUP_Mins])/60 As FUP_P_Hours, ([FUP]*[N_FUP_Mins])/60 As FUP_N_Hours, " &
"([FUP]*[C_FUP_Mins])/60 As FUP_C_Hours, IIf(IsNull([PhysicianHours]),0,[PhysicianHours]) As Q_P_Hours, IIf(IsNull([NurseHours]),0,[NurseHours]) As Q_N_Hours, " &
"IIf(IsNull([CSSHours]),0,[CSSHours]) As Q_C_Hours, IIf(Format([WhatDay],'ddd')='Sat' Or Format([WhatDay],'ddd')='Sun',0,([PermPhysHours]-(([PermPhysHours]/100)*[SickALPercent]))) AS Perm_P_Hours, " &
"IIf(Format([WhatDay],'ddd')='Sat' Or Format([WhatDay],'ddd')='Sun',0,([PermNurseHours]-(([PermNurseHours]/100)*[SickALPercent]))) AS Perm_N_Hours, " &
"IIf(Format([WhatDay],'ddd')='Sat' Or Format([WhatDay],'ddd')='Sun',0,([PermCSSHours]-(([PermCSSHours]/100)*[SickALPercent]))) AS Perm_C_Hours " &
"FROM " &
"hSite_Perm INNER JOIN ((" & ToolLink(Source) & " AS ToolLink " &
"INNER JOIN hSite_Assumptions On ToolLink.hSiteID = hSite_Assumptions.ID) LEFT JOIN hSite_Q_Assumptions On ToolLink.QRisk = hSite_Q_Assumptions.Level) On hSite_Perm.ID = ToolLink.hSite_Perm_ID " &
"ORDER BY " &
"ToolLink.WhatDay" &
")"
    End Function

    Public Function Hours_Per_Month(Source As String) As String
        Return "(" &
"SELECT " &
"Format([WhatDay],'mmm-yyyy') AS WhatMonth, [Hours Per Day].FinanceID, Sum([Hours Per Day].SSS_P_Hours) AS SumOfSSS_P_Hours, " &
"Sum([Hours Per Day].SSS_N_Hours) AS SumOfSSS_N_Hours, Sum([Hours Per Day].SSS_C_Hours) AS SumOfSSS_C_Hours, " &
"Sum([Hours Per Day].PSP_P_Hours) AS SumOfPSP_P_Hours, Sum([Hours Per Day].PSP_N_Hours) AS SumOfPSP_N_Hours, " &
"Sum([Hours Per Day].PSP_C_Hours) AS SumOfPSP_C_Hours, Sum([Hours Per Day].PSP2_P_Hours) AS SumOfPSP2_P_Hours, " &
"Sum([Hours Per Day].PSP2_N_Hours) AS SumOfPSP2_N_Hours, Sum([Hours Per Day].PSP2_C_Hours) AS SumOfPSP2_C_Hours, " &
"Sum([Hours Per Day].FUP_P_Hours) AS SumOfFUP_P_Hours, Sum([Hours Per Day].FUP_N_Hours) AS SumOfFUP_N_Hours, " &
"Sum([Hours Per Day].FUP_C_Hours) AS SumOfFUP_C_Hours, Sum([Hours Per Day].Q_P_Hours) AS SumOfQ_P_Hours, " &
"Sum([Hours Per Day].Q_N_Hours) AS SumOfQ_N_Hours, Sum([Hours Per Day].Q_C_Hours) AS SumOfQ_C_Hours, " &
"Sum([Hours Per Day].Perm_P_Hours) AS SumOfPerm_P_Hours, Sum([Hours Per Day].Perm_N_Hours) AS SumOfPerm_N_Hours, Sum([Hours Per Day].Perm_C_Hours) AS SumOfPerm_C_Hours " &
"FROM " &
Hours_Per_Day(Source) & " AS [Hours Per Day] " &
"GROUP BY " &
"Format([WhatDay],'mmm-yyyy'), [Hours Per Day].FinanceID, CDate(Format([WhatDay],'mmm-yyyy')) " &
"ORDER BY " &
"CDate(Format([WhatDay],'mmm-yyyy')) " &
")"
    End Function

    Public Function Hours_Per_Week(Source As String) As String
        Return "(" &
"SELECT " &
"Min([Hours Per Day].WhatDay) AS MinOfWhatDay, [Hours Per Day].FinanceID, Sum([Hours Per Day].SSS_P_Hours) As SumOfSSS_P_Hours, " &
"Sum([Hours Per Day].SSS_N_Hours) As SumOfSSS_N_Hours, Sum([Hours Per Day].SSS_C_Hours) As SumOfSSS_C_Hours, Sum([Hours Per Day].PSP_P_Hours) As SumOfPSP_P_Hours, " &
"Sum([Hours Per Day].PSP_N_Hours) As SumOfPSP_N_Hours, Sum([Hours Per Day].PSP_C_Hours) As SumOfPSP_C_Hours, Sum([Hours Per Day].PSP2_P_Hours) As SumOfPSP2_P_Hours, " &
"Sum([Hours Per Day].PSP2_N_Hours) As SumOfPSP2_N_Hours, Sum([Hours Per Day].PSP2_C_Hours) As SumOfPSP2_C_Hours, Sum([Hours Per Day].FUP_P_Hours) As SumOfFUP_P_Hours, " &
"Sum([Hours Per Day].FUP_N_Hours) As SumOfFUP_N_Hours, Sum([Hours Per Day].FUP_C_Hours) As SumOfFUP_C_Hours, Sum([Hours Per Day].Q_P_Hours) As SumOfQ_P_Hours, " &
"Sum([Hours Per Day].Q_N_Hours) As SumOfQ_N_Hours, Sum([Hours Per Day].Q_C_Hours) As SumOfQ_C_Hours, Sum([Hours Per Day].Perm_P_Hours) As SumOfPerm_P_Hours, " &
"Sum([Hours Per Day].Perm_N_Hours) As SumOfPerm_N_Hours, Sum([Hours Per Day].Perm_C_Hours) As SumOfPerm_C_Hours " &
"FROM " &
Hours_Per_Day(Source) & " AS [Hours Per Day] " &
"GROUP BY " &
"[Hours Per Day].FinanceID, DatePart('ww',[WhatDay],2,1) & ' ' & DatePart('yyyy',[WhatDay],2,1) " &
"ORDER BY " &
"Min([Hours Per Day].WhatDay)" &
")"
    End Function

    Public Function Hours_Per_Month_Surplus(Source As String) As String
        Return "(" &
"SELECT " &
"[Hours Per Month].WhatMonth, ([sumofsss_P_Hours]+[SumofPSP_P_Hours]+[SumofFUP_P_Hours]+[SumofQ_P_Hours])-[Sumofperm_P_Hours] AS Hours_P_Need, " &
"([sumofsss_N_Hours]+[SumofPSP_N_Hours]+[SumofFUP_N_Hours]+[SumofQ_N_Hours])-[Sumofperm_N_Hours] AS Hours_N_Need, " &
"([sumofsss_C_Hours]+[SumofPSP_C_Hours]+[SumofFUP_C_Hours]+[SumofQ_C_Hours])-[Sumofperm_C_Hours] AS Hours_C_Need, " &
"([sumofsss_P_Hours]+[SumofPSP_P_Hours]+[SumofFUP_P_Hours]+[SumofQ_P_Hours]-[Sumofperm_P_Hours])*[P_GBP_Hour] AS P_GBP, " &
"([sumofsss_N_Hours]+[SumofPSP_N_Hours]+[SumofFUP_N_Hours]+[SumofQ_N_Hours]-[Sumofperm_N_Hours])*[N_GBP_Hour] AS N_GBP, " &
"([sumofsss_C_Hours]+[SumofPSP_C_Hours]+[SumofFUP_C_Hours]+[SumofQ_C_Hours]-[Sumofperm_C_Hours])*[C_GBP_Hour] AS C_GBP, CDate([Whatmonth]) AS DateField " &
"FROM " &
Hours_Per_Month(Source) & " AS [Hours Per Month] INNER JOIN Finance_Assumptions ON [Hours Per Month].FinanceID = Finance_Assumptions.ID " &
"ORDER BY " &
"CDate([Whatmonth])" &
")"
    End Function

    Public Function Hours_Per_Week_Surplus(Source As String) As String
        Return "(" &
"SELECT " &
"[Hours Per Week].MinOfWhatDay, ([sumofsss_P_Hours]+[SumofPSP_P_Hours]+[SumofFUP_P_Hours]+[SumofQ_P_Hours])-[Sumofperm_P_Hours] AS Hours_P_Need, " &
"([sumofsss_N_Hours]+[SumofPSP_N_Hours]+[SumofFUP_N_Hours]+[SumofQ_N_Hours])-[Sumofperm_N_Hours] AS Hours_N_Need, " &
"([sumofsss_C_Hours]+[SumofPSP_C_Hours]+[SumofFUP_C_Hours]+[SumofQ_C_Hours])-[Sumofperm_C_Hours] AS Hours_C_Need " &
"FROM " &
Hours_Per_Week(Source) & " AS [Hours Per Week]" &
")"

    End Function

    Public Function hSite_Bank_Per_Month(Source As String) As String
        Return "(" &
"SELECT " &
"[Hours Per Month Surplus].WhatMonth, [Hours Per Month Surplus].DateField, " &
"IIf([Hours_P_Need]-Int([Hours_P_Need])=0,[Hours_P_Need],IIf(Int([Hours_P_Need])>0,Int([Hours_P_Need])+1,Int([Hours_P_Need]))) AS P_Hours, " &
"IIf([Hours_N_Need]-Int([Hours_N_Need])=0,[Hours_N_Need],IIf(Int([Hours_N_Need])>0,Int([Hours_N_Need])+1,Int([Hours_N_Need]))) AS N_Hours, " &
"IIf([Hours_C_Need]-Int([Hours_C_Need])=0,[Hours_C_Need],IIf(Int([Hours_C_Need])>0,Int([Hours_C_Need])+1,Int([Hours_C_Need]))) AS C_Hours, " &
"IIf([P_GBP]-Int([P_GBP])=0,[P_GBP],IIf(Int([P_GBP])>0,Int([P_GBP])+1,Int([P_GBP]))) AS P_Price, IIf([N_GBP]-Int([N_GBP])=0,[N_GBP],IIf(Int([N_GBP])>0,Int([N_GBP])+1,Int([N_GBP]))) AS N_Price, " &
"IIf([C_GBP]-Int([C_GBP])=0,[C_GBP],IIf(Int([C_GBP])>0,Int([C_GBP])+1,Int([C_GBP]))) AS C_Price " &
"FROM " &
Hours_Per_Month_Surplus(Source) & " AS [Hours Per Month Surplus]" &
")"

    End Function

    Public Function hSite_Bank_Per_Week(Source As String) As String
        Return "(" &
"SELECT " &
"[Hours Per Week Surplus].MinOfWhatDay, IIf([Hours_P_Need]-Int([Hours_P_Need])=0,[Hours_P_Need],IIf(Int([Hours_P_Need])>0,Int([Hours_P_Need])+1,Int([Hours_P_Need]))) AS P_Hours, " &
"IIf([Hours_N_Need]-Int([Hours_N_Need])=0,[Hours_N_Need],IIf(Int([Hours_N_Need])>0,Int([Hours_N_Need])+1,Int([Hours_N_Need]))) AS N_Hours, " &
"IIf([Hours_C_Need]-Int([Hours_C_Need])=0,[Hours_C_Need],IIf(Int([Hours_C_Need])>0,Int([Hours_C_Need])+1,Int([Hours_C_Need]))) AS C_Hours " &
"FROM " &
Hours_Per_Week_Surplus(Source) & " AS [Hours Per Week Surplus]" &
")"
    End Function

    Public Function ToolLink(Source As String) As String

        Return "(" &
"SELECT " &
"Max(hSite_Perm.ID) As hSite_Perm_ID, Max(hSite_Assumptions.ID) As hSiteID, Max(Finance_Assumptions.id) As FinanceID, " &
"Max(hSite_Q_Assumptions.GroupID) As QID, a.WhatDay, a.PSP, a.SSS, a.PSP2, a.Fup, a.QRisk " &
"FROM " &
"hSite_Assumptions, Finance_Assumptions, hSite_Perm, " & Source & " AS a INNER JOIN DaysLink On a.WhatDay = DaysLink.WhatDay, Q_Group INNER JOIN hSite_Q_Assumptions On Q_Group.ID = hSite_Q_Assumptions.GroupID " &
"WHERE " &
"(((Q_Group.Selected)=True) And ((Finance_Assumptions.Selected)=True) And ((hSite_Assumptions.Selected)=True) And ((hSite_Perm.Selected)=True)) " &
"GROUP BY " &
"a.WhatDay, a.PSP, a.SSS, a.PSP2, a.Fup, a.QRisk" &
")"

    End Function

End Module
