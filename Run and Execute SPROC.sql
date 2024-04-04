
CREATE Procedure SP_Total_Contracts_Traded_Report
	@DateFrom datetime2,
	@DateTo datetime2
AS 
BEGIN


  SELECT FileDate, Contract, SUM(ContractsTraded) ContractsTraded, 
  SUM(ContractsTraded) * 100 / SUM(SUM(ContractsTraded)) over (PARTITION BY FileDate) ['Percentage Of Total ContractsTraded']
  FROM DailyMTM 
    WHERE FileDate >= @DateFrom 
    AND FileDate <= @DateTo
  GROUP BY FileDate, Contract
  HAVING SUM(ContractsTraded) > 0
  
  ORDER BY FileDate, Contract
END 
GO

GO 
BEGIN 
--DECLARE @DateFrom datetime2 = '2021-01-04';
--DECLARE @DateTo datetime2 = '2021-01-05';

EXEC SP_Total_Contracts_Traded_Report 
@DateFrom = N'2021-01-04', @DateTo = N'2021-01-05'

END 
GO