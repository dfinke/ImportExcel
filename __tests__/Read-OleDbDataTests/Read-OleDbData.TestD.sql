select top 1
    'All A1s' as [A1],
    F1 as [Sheet1],
    (select F1 FROM [sheet2$a1:a1]) as [Sheet2],
    (select F1 FROM [sheet3$a1:a1]) as [Sheet3],
    (select F1 FROM [sheet4$a1:a1]) as [Sheet4],
    (select F1 FROM [sheet5$a1:a1]) as [Sheet5],
    (select F1 FROM [sheet6$a1:a1]) as [Sheet6],
    (select F1 FROM [sheet7$a1:a1]) as [Sheet7]
FROM [sheet1$a1:a1]