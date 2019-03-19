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
       T5."Price" AS "含税采购价格",
       T5."Currency" AS "采购货币",
       case when T5."Currency"='USD' then T5."Price" * 6.7 else T5."Price" end AS "含税人民币价格",
       T1."AvgPrice" - case when T5."Currency"='USD' then T5."Price" * 6.7 else T5."Price" end / 1.16 AS "价格差异",
       (T1."AvgPrice" - case when T5."Currency"='USD' then T5."Price" * 6.7 else T5."Price" end / 1.16) * T1."OnHand" AS "总成本差异",
       (T1."AvgPrice" - case when T5."Currency"='USD' then T5."Price" * 6.7 else T5."Price" end / 1.16) / T1."AvgPrice" AS "价格差异比率"
FROM OITM T0
  INNER JOIN OITW T1 ON T0."ItemCode" = T1."ItemCode"
  INNER JOIN OWHS T2 ON T1."WhsCode" = T2."WhsCode"
  INNER JOIN OITB T3 ON T0."ItmsGrpCod" = T3."ItmsGrpCod"
  LEFT JOIN OBCD T4 ON T0."ItemCode" = T4."ItemCode"
  LEFT JOIN ITM1 T5 on T1."ItemCode" = T5."ItemCode" and T5."PriceList"='1'
WHERE T1."AvgPrice" <> '0'
  and T2."WhsName" like '%生产%'
  and  T0."ItmsGrpCod"<>'110'
  and  T1."WhsCode" like '%109%';
