-- Query: qryCreateKPI_Month_Results
-- DAO query type: 80

SELECT qryKPI_Results_Monthly.ProjName, qryKPI_Results_Monthly.WorkOrdr, qryKPI_Results_Monthly.SmplPerE, qryKPI_Results_Monthly.Month, qryKPI_Results_Monthly.CountOfSampleId, qryKPI_Results_Monthly.SampReleaseDate, qryKPI_Results_Monthly.LastSampRecDate, qryKPI_Results_Monthly.ReportDueDate, qryKPI_Results_Monthly.Category, qryKPI_Results_Monthly.Late, qryKPI_Results_Monthly.CountOfLate, qryKPI_Results_Monthly.Year, qryKPI_Results_Monthly.monthtxt, qryKPI_Results_Monthly.YearNumber INTO KPI_Results_Monthly
FROM qryKPI_Results_Monthly;


