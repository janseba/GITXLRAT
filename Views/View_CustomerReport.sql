SELECT Country
	,PlanVersion
	,Period
	,Cint(Right(Period, 2)) AS [Month]
	,IIf(ISNULL(a.EUProductHierarchy), "NA", a.EUProductHierarchy) AS EUProductHierarchy
	,IIf(ISNULL(a.Brand), "NA", a.Brand) AS Brand
	,IIf(ISNULL(a.Prdha2), "NA", a.Prdha2) AS Prdha2
	,IIf(ISNULL(a.ProfitCenter), "NA", a.ProfitCenter) AS ProfitCenter
	,IIf(ISNULL(a.PlanningCustomer), "NA", a.PlanningCustomer) AS PlanningCustomer
	,IIf(ISNULL(a.ConditionCustomer), "NA", a.ConditionCustomer) AS ConditionCustomer
	,IIf(ISNULL(a.CustomerName), "NA", a.CustomerName) AS CustomerName
	,IIf(ISNULL(a.ReportingCategory), "NA", a.ReportingCategory) AS ReportingCategory
	,IIf(ISNULL(a.Prdha3), "NA", a.Prdha3) AS Prdha3
	,IIf(ISNULL(a.Prdha4), "NA", a.Prdha4) AS Prdha4
	,IIf(ISNULL(a.Prdha1), "NA", a.Prdha1) AS Prdha1
	,SUM(a.Volume) AS Volume
	,SUM(a.Drinks) AS Drinks
	,SUM(a.tDiscs) AS tDiscs
	,SUM(a.GrossSalesValueInclCT) AS GrossSalesValueInclCT
	,SUM(a.GrossSalesValueExclCT) AS GrossSalesValueExclCT
	,SUM(a.PPR_LPA) AS PPR_LPA
	,SUM(a.OnInvoiceConditions) AS OnInvoiceConditions
	,SUM(a.NISInclCT) AS NISInclCT
	,SUM(a.PPR) AS PPR
	,SUM(a.TPR) AS TPR
	,SUM(a.BDF) AS BDF
	,SUM(a.BMC) AS BMC
	,SUM(a.TradeSpend) AS TradeSpend
	,SUM(a.NOSInclCT) AS NOSInclCT
	,SUM(a.NOSExclCT) AS NOSExclCT
	,SUM(a.TotMB) AS MB
	,SUM(a.mbadjustments) AS mbadjustments
	,SUM(a.TotGreendot) AS Greendot
	,SUM(a.TotDisplayCosts) AS DisplayCosts
	,SUM(a.CostOfGoodsExclCT) AS CostOfGoodsExclCT
	,SUM(a.TotCoffeeTax) AS CoffeeTax
	,SUM(a.CostOfGoodsInclCT) AS CostOfGoodsInclCT
	,SUM(a.GP) AS GP
	,SUM(a.advworkingmedia) AS advworkingmedia
	,SUM(a.advnonworkingmedia) AS advnonworkingmedia
	,SUM(a.Advertising) AS Advertising
	,SUM(a.brewersupport) AS brewersupport
	,SUM(a.promotionother) AS promotion
	,SUM(a.Promotion) AS TotPromotion
	,SUM(a.TotalAP) AS TotalAP
	,SUM(a.MarketingCM) AS MarketingCM
	,SUM(a.WD) AS WD
	,SUM(a.CM) AS CM
	,SUM(a.OSA) AS OSA
	,SUM(a.Ebit) AS Ebit
	,SUM(a.Discount5) AS Discount5
	,SUM(a.ListingFees) AS ListingFees
	,SUM(a.FixedAmounts) AS FixedAmounts
FROM View_Facts AS a
GROUP BY Country
	,PlanVersion
	,Period
	,EUProductHierarchy
	,Brand
	,Prdha2
	,ProfitCenter
	,PlanningCustomer
	,ConditionCustomer
	,CustomerName
	,ReportingCategory
	,Prdha3
	,Prdha4
	,Prdha1;
