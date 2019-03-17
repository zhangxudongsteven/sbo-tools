select
	T0."Code" AS "物料编码",
	T2."ItemName" AS "物料名称",
    T2."SuppCatNum" AS "原厂料号",
    case when T2."PrcrmntMtd"='B' then '购买' when T2."PrcrmntMtd"='M' then '生产' else '未知' end AS "补给方式",
    T2."CardCode" AS "首选供应商",
    T3."Name" AS "物料分类二",
    T4."Name" AS "物料分类三",
    T2."BuyUnitMsr" AS "计量单位",
	T0."Father" AS "上级物料编码",
	T1."ItemName" AS "上级物料名称",
	T0."Quantity" AS "总消耗量",
	T0."U_StandLoss" AS "损耗量",
	T2."U_MaterialLoss" AS "标准损耗率",
	T5."Price" AS "采购价格",
	T5."Currency" AS "采购货币"
from ITT1 T0
	left join OITM T1 on T0."Father" = T1."ItemCode"
	left join OITM T2 on T0."Code" = T2."ItemCode"
	left join "@U_CIICL2" T3 on T2."U_Class2" = T3."Code"
	left join "@U_CIICL3" T4 on T2."U_Class3" = T4."Code"
	left join ITM1 T5 on T0."Code" = T5."ItemCode" and T5."PriceList"='1'
order by T0."Father"