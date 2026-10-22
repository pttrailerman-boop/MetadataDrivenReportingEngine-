-- Query: qryInsertInto_SorKPI_MonthlyReport
-- DAO query type: 80

SELECT [Year] & Left([Yearnumber],2) AS SpecialSort, [monthtxt] & " " & [Year] AS Month_Year, Count(KPI_Results_Monthly.WorkOrdr) AS CountOfWorkOrdr INTO tblSorKPI_MonthlyReport
FROM KPI_Results_Monthly
GROUP BY [Year] & Left([Yearnumber],2), [monthtxt] & " " & [Year]
ORDER BY [Year] & Left([Yearnumber],2);


