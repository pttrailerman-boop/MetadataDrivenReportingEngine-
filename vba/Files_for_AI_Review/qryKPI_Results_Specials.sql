-- Query: qryKPI_Results_Specials
-- DAO query type: 0

SELECT qryKPI_Results.ProjName, qryKPI_Results.WorkOrdr, qryKPI_Results.SmplPerE, qryKPI_Results.Month, qryKPI_Results.CountOfSampleId, qryKPI_Results.SampReleaseDate, qryKPI_Results.LastSampRecDate, qryKPI_Results.ReportDueDate, qryKPI_Results.Category, qryKPI_Results.Late, qryKPI_Results.CountOfLate, qryKPI_Results.Year, qryKPI_Results.monthtxt, qryKPI_Results.YearNumber, qryBegin_Yr.Begin_Yr, qryEnd_yr.End_Yr
FROM qryKPI_Results, qryBegin_Yr, qryEnd_yr
WHERE (((qryKPI_Results.Category)="Special"))
ORDER BY qryKPI_Results.SmplPerE DESC;


