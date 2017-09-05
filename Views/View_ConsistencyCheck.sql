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
UNION ALL 
SELECT a.planversion, 
       a.sku, 
       'Missing Weight In KG' 
FROM   tblfacts AS a 
       LEFT JOIN tblsku AS b 
              ON a.sku = b.sku 
WHERE  ( b.sku IS NULL 
          OR b.weightinkg IS NULL 
          OR b.weightinkg = 0 ) 
       AND forecast = 'yes' 
GROUP  BY a.planversion, 
          a.sku 
HAVING Sum(a.volume) <> 0 
UNION ALL 
SELECT a.planversion, 
       a.sku, 
       'Missing SalesConditionLevel' 
FROM   tblfacts AS a 
       LEFT JOIN tblsku AS b 
              ON a.sku = b.sku 
WHERE  a.forecast = 'yes' 
       AND b.salesconditionlevel IS NULL 
GROUP  BY a.planversion, 
          a.sku 
HAVING Sum(a.volume) <> 0 
UNION ALL 
SELECT fact.planversion, 
       fact.sku + " | " + fact.period, 
       'Missing FAP' 
FROM   (SELECT a.planversion, 
               a.sku, 
               a.period, 
               b.salesconditionlevel, 
               a.forecast, 
               a.volume 
        FROM   tblfacts AS a 
               INNER JOIN tblsku AS B 
                       ON a.sku = b.sku) AS fact 
       LEFT JOIN tblFap AS fap 
              ON fact.planversion = fap.planversion 
                 AND fact.salesconditionlevel = fap.salescondition 
                 AND fact.period = fap.period 
WHERE  fap.planversion IS NULL 
       AND fact.forecast = 'yes' 
GROUP  BY fact.planversion, 
          fact.sku, 
          fact.period
HAVING Sum(fact.volume) <> 0