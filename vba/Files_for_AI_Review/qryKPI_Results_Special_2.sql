-- Query: qryKPI_Results_Special_2
-- DAO query type: 80

SELECT qryKPI_Results_Specials.ProjName, qryKPI_Results_Specials.WorkOrdr, qryKPI_Results_Specials.SmplPerE, qryKPI_Results_Specials.Month, qryKPI_Results_Specials.SampReleaseDate, qryKPI_Results_Specials.LastSampRecDate, qryKPI_Results_Specials.ReportDueDate, qryKPI_Results_Specials.Category, qryKPI_Results_Specials.Late, qryKPI_Results_Specials.CountOfLate, 1 AS countofWorkOrder, ([Countofworkorder] - [countofLate]) / [Countofworkorder] AS PercentWoOnTime, qryKPI_Results_Specials.CountOfSampleId, IIf([countoflate] > 0, [Countofsampleid], 0) AS CountofLateSamples, ([countofsampleid] - [countoflatesamples]) / [countofsampleid] AS PercentSamplesOnTime, qryKPI_Results_Specials.Year, qryKPI_Results_Specials.monthtxt, qryKPI_Results_Specials.YearNumber, [Begin_Yr] & " To " & [End_yr] AS Year_Range INTO KPI_Results_Special
FROM qryKPI_Results_Specials
ORDER BY qryKPI_Results_Specials.SmplPerE;


