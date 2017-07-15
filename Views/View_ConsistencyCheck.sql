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
       fact.sku + " | " + fact.customer + " | " 
       + fact.period, 
       'Missing PPP' 
FROM   (SELECT a.planversion, 
               a.customer, 
               a.sku, 
               a.period, 
               b.salesconditionlevel, 
               a.forecast, 
               a.volume 
        FROM   tblfacts AS a 
               INNER JOIN tblsku AS B 
                       ON a.sku = b.sku) AS fact 
       LEFT JOIN tblppp AS ppp 
              ON fact.planversion = ppp.planversion 
                 AND fact.customer = ppp.customer 
                 AND fact.salesconditionlevel = ppp.salesconditionlevel 
                 AND fact.period = ppp.period 
WHERE  ppp.planversion IS NULL 
       AND fact.forecast = 'yes' 
GROUP  BY fact.planversion, 
          fact.sku, 
          fact.period, 
          fact.customer 
HAVING Sum(fact.volume) <> 0   