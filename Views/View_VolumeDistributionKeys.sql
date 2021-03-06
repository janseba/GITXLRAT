SELECT a.PlanVersion,
       b.AlternativeSKU,
       a.Customer,
       IIf(ISNULL(b.SalesPlanning),'NA',b.SalesPlanning) AS SalesPlanning,
       SUM(a.Volume) AS Volume,
       SUM(c.Volume) AS TotVol,
       SUM(IIf(c.Volume=0,0,a.Volume/c.Volume)) AS DistrKey
FROM (tblReferenceVolume AS a
      INNER JOIN tblSKU AS b ON a.SKU = b.SKU)
INNER JOIN View_ReferenceVolumeOnPlanningLevel AS c ON (IIf(ISNULL(a.Customer),'NA',a.Customer)=c.Customer)
AND (IIf(ISNULL(b.SalesPlanning),'NA',b.SalesPlanning)=c.SalesPlanning)
AND (a.PlanVersion = c.PlanVersion)
GROUP BY a.PlanVersion, b.AlternativeSKU, a.Customer, b.SalesPlanning

