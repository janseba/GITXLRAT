SELECT source                                         AS Planversion, 
       alternativesku                                 AS ErrorObject, 
       'Alternative SKU has more than one active SKU' AS ErrorType 
FROM   tblsku, 
       sources 
WHERE  active = 'yes' 
GROUP  BY source, 
          alternativesku 
HAVING Count(alternativesku) > 1 
UNION ALL 
SELECT a.planversion, 
       a.sku, 
       'SKU is missing' 
FROM   tblfacts AS a 
       LEFT JOIN tblsku AS b 
              ON a.sku = b.sku 
WHERE  b.sku IS NULL 
GROUP  BY a.planversion, 
          a.sku 
HAVING Sum(a.volume) <> 0   