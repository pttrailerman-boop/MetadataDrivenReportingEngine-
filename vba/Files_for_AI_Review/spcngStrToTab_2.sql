-- Query: spcngStrToTab_2
-- DAO query type: 0

SELECT tblReportSequence.*
FROM tblReportSequence
WHERE (((tblReportSequence.Tab) Is Not Null) And ((tblReportSequence.Report) = "Monthly_Report.XLSX"))
ORDER BY tblReportSequence.[Spacing Order], tblReportSequence.Sort;


