SELECT a.PlanVersion,
       a.SKU,
       a.Customer,
       IIf(ISNULL(b.SalesPlanning),'NA',b.SalesPlanning) AS SalesPlanning,
       a.Volume,
       c.Volume AS TotVol,
       IIf(c.Volume=0,0,a.Volume/c.Volume) AS DistrKey
FROM (tblReferenceVolume AS a
      INNER JOIN tblSKU AS b ON a.SKU = b.SKU)
INNER JOIN View_ReferenceVolumeOnPlanningLevel AS c ON (IIf(ISNULL(a.Customer),'NA',a.Customer)=c.Customer)
AND (IIf(ISNULL(b.SalesPlanning),'NA',b.SalesPlanning)=c.SalesPlanning)
AND (a.PlanVersion = c.PlanVersion);

