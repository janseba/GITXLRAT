SELECT a.PlanVersion,
       a.Period,
       a.SourceType,
       a.Forecast,
       a.SKU,
       b.Description AS SKUDescription,
       b.Prdha4,
       b.Prdha3,
       b.Prdha2,
       b.Prdha1,
       b.PromotionalSKU,
       b.ProfitCenter,
       b.Brand,
       b.BridgeHierarchy,
       b.ReportingCategory,
       a.Customer,
       c.CustomerName,
       c.PlanningCustomer,
       a.PromoNonPromo,
       a.OnOffInvoice,
       SUM(a.Volume) AS Volume,
       SUM(a.VolPromo) AS VolPromo,
       SUM(a.VolNonPromo) AS VolNonPromo,
       SUM(a.Ebit) AS Ebit,
       SUM(a.OSA) AS OSA,
       SUM(a.CM) AS CM,
       SUM(a.WD) AS WD,
       SUM(a.MarketingCM) AS MarketingCM,
       SUM(a.GP) AS GP,
       SUM(a.TotalAP) AS TotalAP,
       SUM(a.Advertising) AS Advertising,
       SUM(a.Promotion) AS Promotion,
       SUM(a.NOSInclCT) AS NOS,
       SUM(a.GrossSalesValueInclCT) AS GOS,
       SUM(a.TradeSpend) AS TradeSpend,
       SUM(a.PPR_LPA) AS LPA,
       SUM(a.PPR) AS PPR,
       SUM(a.TPR) AS TPR,
       SUM(a.OnInvoiceConditions) AS OnInvoiceConditions,
       SUM(a.GrossSalesValueInclCT + a.PPR_LPA + a.OnInvoiceConditions) AS NIS,
       SUM(a.BDF) AS BDF,
       SUM(a.BMC) AS BMC,
       SUM(a.CostOfGoodsExclCT) AS COGS,
       SUM(a.TotDisplayCosts) AS DisplayCosts,
       SUM(a.TotMB) AS MB,
       SUM(a.TotGreendot) AS Greendot
FROM (view_facts AS a
      LEFT JOIN tblSKU AS b ON a.SKU = b.SKU)
LEFT JOIN tblCustomer AS c ON a.Customer = c.Customer
GROUP BY a.PlanVersion,
         a.Period,
         a.SourceType,
         a.Forecast,
         a.SKU,
         b.Description,
         b.Prdha4,
         b.Prdha3,
         b.Prdha2,
         b.Prdha1,
         b.PromotionalSKU,
         b.ProfitCenter,
         b.Brand,
         b.BridgeHierarchy,
         b.ReportingCategory,
         a.Customer,
         c.CustomerName,
         c.PlanningCustomer,
         a.PromoNonPromo,
         a.OnOffInvoice