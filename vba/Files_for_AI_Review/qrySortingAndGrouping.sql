-- Query: qrySortingAndGrouping
-- DAO query type: 0

SELECT *
FROM tbl_temp_sorting_grouping
WHERE tbl_temp_sorting_grouping.LEVEL = "tblSorKPI_SpecialReport"
ORDER BY tbl_temp_sorting_grouping.Sort1 DESC;


