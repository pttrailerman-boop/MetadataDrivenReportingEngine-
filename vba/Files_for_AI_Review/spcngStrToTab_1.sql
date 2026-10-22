-- Query: spcngStrToTab_1
-- DAO query type: 0

SELECT tblReportSequence.*
FROM tblReportSequence
WHERE (((tblReportSequence.Spacing) Is Not Null) And ((tblReportSequence.Report) = "Monthly_Report.XLSX"))
ORDER BY tblReportSequence.[Spacing Order], tblReportSequence.Sort;


