-- Query: qryInsertInto_SorKPI_SpeicalReport
-- DAO query type: 80

SELECT [Year] & Left([Yearnumber],2) AS SpecialSort, [monthtxt] & " " & [Year] AS Month_Year, Count(KPI_Results_Special.WorkOrdr) AS CountOfWorkOrdr INTO tblSorKPI_SpecialReport
FROM KPI_Results_Special
GROUP BY [Year] & Left([Yearnumber],2), [monthtxt] & " " & [Year]
ORDER BY [Year] & Left([Yearnumber],2);


