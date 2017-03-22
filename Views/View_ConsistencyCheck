SELECT SOURCE AS Planversion,
                 AlternativeSKU AS ErrorObject,
                 'Alternative SKU has more than one active SKU' AS ErrorType
FROM tblSKU,
     Sources
WHERE Active ='yes'
GROUP BY SOURCE,
         AlternativeSKU HAVING COUNT(AlternativeSKU) > 1