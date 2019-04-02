SELECT
  T0."ItemCode",
  T0."ItemName",
  T4."BcdCode",
  T0."ItmsGrpCod",
  T3."ItmsGrpNam",
  T1."WhsCode",
  T2."WhsName",
  T1."OnHand",
  T1."AvgPrice" AS "加权平均成本",
  T5."Price" AS "未税采购价格",
  T5."Currency" AS "采购货币",
  CASE WHEN T5."Currency"='USD' THEN T5."Price" * 6.7 ELSE T5."Price" END AS "未税人民币价格",
  (T1."AvgPrice" - CASE WHEN T5."Currency"='USD' THEN T5."Price" * 6.7 ELSE T5."Price" END) AS "价格差异",
  (T1."AvgPrice" - CASE WHEN T5."Currency"='USD' THEN T5."Price" * 6.7 ELSE T5."Price" END) * T1."OnHand" AS "总成本差异",
  (T1."AvgPrice" - CASE WHEN T5."Currency"='USD' THEN T5."Price" * 6.7 ELSE T5."Price" END) / T1."AvgPrice" AS "价格差异比率"
FROM OITM T0
  INNER JOIN OITW T1 ON T0."ItemCode" = T1."ItemCode"
  INNER JOIN OWHS T2 ON T1."WhsCode" = T2."WhsCode"
  INNER JOIN OITB T3 ON T0."ItmsGrpCod" = T3."ItmsGrpCod"
  LEFT JOIN OBCD T4 ON T0."ItemCode" = T4."ItemCode"
  LEFT JOIN ITM1 T5 ON T1."ItemCode" = T5."ItemCode" AND T5."PriceList"='4'
WHERE T1."AvgPrice" <> '0'
  AND T2."WhsName" LIKE '%生产%'
  AND T0."ItmsGrpCod" <> '110'
  AND T1."WhsCode" LIKE '%109%';
