-- Query: spcngStrToTab
-- DAO query type: 0

SELECT spcngStrToTab_1.id, spcngStrToTab_1.Report, spcngStrToTab_1.Sort, spcngStrToTab_1.version, spcngStrToTab_1.QueryName, spcngStrToTab_1.FinalOutput, spcngStrToTab_2.Tab, spcngStrToTab_1.[Tab Order], spcngStrToTab_1.[isTable?], spcngStrToTab_1.Sql, spcngStrToTab_1.Spacing, spcngStrToTab_1.[Spacing Order], spcngStrToTab_1.Exec, spcngStrToTab_1.[HasTotals], spcngStrToTab_1.[Over_ride_head1], spcngStrToTab_1.[Over_ride_foot1], spcngStrToTab_1.Notes
FROM spcngStrToTab_1 INNER JOIN spcngStrToTab_2 ON spcngStrToTab_1.Spacing = spcngStrToTab_2.Tab
ORDER BY spcngStrToTab_1.[Spacing Order];


