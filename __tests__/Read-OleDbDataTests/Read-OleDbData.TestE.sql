select top 1
    'All A1s Start from Sheet1' as [A1],
    F1 as [Sheet1],
    (select F1 FROM [sheet2$a1:a1]) as [Sheet2],
    (select F1 FROM [sheet3$a1:a1]) as [Sheet3],
    (select F1 FROM [sheet4$a1:a1]) as [Sheet4]
FROM [sheet1$a1:a1]
UNION ALL
select top 1
    'All A1s Start from Sheet2' as [A1],
    (select F1 FROM [sheet1$a1:a1]) as [Sheet1],
    F1 as [Sheet2],
    (select F1 FROM [sheet3$a1:a1]) as [Sheet3],
    (select F1 FROM [sheet4$a1:a1]) as [Sheet4]
FROM [sheet2$a1:a1]
UNION ALL
select top 1
    'All A1s Start from Sheet3' as [A1],
    (select F1 FROM [sheet1$a1:a1]) as [Sheet1],
    (select F1 FROM [sheet2$a1:a1]) as [Sheet2],
    F1 as [Sheet3],
    (select F1 FROM [sheet4$a1:a1]) as [Sheet4]
FROM [sheet3$a1:a1]
UNION ALL
select top 1
    'All A1s Start from Sheet4' as [A1],
    (select F1 FROM [sheet1$a1:a1]) as [Sheet1],
    (select F1 FROM [sheet2$a1:a1]) as [Sheet2],
    (select F1 FROM [sheet3$a1:a1]) as [Sheet3],
    F1 as [Sheet4]
FROM [sheet4$a1:a1]
