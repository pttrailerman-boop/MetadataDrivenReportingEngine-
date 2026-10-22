-- Query: qryKPI_Results_Monthly_2
-- DAO query type: 80

SELECT qryKPI_Results_Monthly.ProjName, qryKPI_Results_Monthly.WorkOrdr, qryKPI_Results_Monthly.SmplPerE, qryKPI_Results_Monthly.Month, qryKPI_Results_Monthly.SampReleaseDate, qryKPI_Results_Monthly.LastSampRecDate, qryKPI_Results_Monthly.ReportDueDate, qryKPI_Results_Monthly.Category, qryKPI_Results_Monthly.Late, qryKPI_Results_Monthly.CountOfLate, 1 AS countofWorkOrder, ([Countofworkorder] - [countofLate]) / [Countofworkorder] AS PercentWoOnTime, qryKPI_Results_Monthly.CountOfSampleId, IIf([countoflate] > 0, [Countofsampleid], 0) AS CountofLateSamples, ([countofsampleid] - [countoflatesamples]) / [countofsampleid] AS PercentSamplesOnTime, qryKPI_Results_Monthly.Year, qryKPI_Results_Monthly.monthtxt, qryKPI_Results_Monthly.YearNumber, [Begin_Yr] & " To " & [End_yr] AS Year_Range INTO KPI_Results_Monthly
FROM qryKPI_Results_Monthly
ORDER BY qryKPI_Results_Monthly.SmplPerE;


