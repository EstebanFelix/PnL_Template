USE [SQLEXPRESS]
GO

/****** Object:  View [dbo].[Budget]    Script Date: 6/25/2023 2:16:16 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO




CREATE VIEW [dbo].[Budget]

AS

SELECT
	[Budget Code],
	[Budget Version],
	[Company Code],
	[Company Name],
	[Fiscal Year],
	[GL Account],
	[GL Account Description],
	[Reporting Code],
	[Fiscal Period],
	[Location Dim 3],
	[Ledger Posting Amount],
	[CONTROLLABLE PROFIT],
	[CONTROLLABLE PROFIT - IHOP],
	[CONTROLLABLE PROFIT - Papa John's],
	[NET EBITDA - Landlord],
	[NET EBITDA - Landlord PJ],
	[NET PROFIT (LOSS)],
	[NET PROFIT (LOSS) - FRE],
	[NET PROFIT (LOSS) - IHOP],
	[NET PROFIT (LOSS) - Papa John's],
	[NET PROFIT (LOSS) - Detail],
	[NET PROFIT (LOSS) - Detail FRE],
	[NET PROFIT (LOSS) - Detail Papa John's]
FROM

	(
	SELECT
		b.BUD_0 AS [Budget Code],
		b.VER_0 AS [Budget Version],
		b.CPY_0 AS [Company Code],
		c.CPYNAM_0 AS [Company Name],
		b.FIY_0 AS [Fiscal Year],
		b.ACC_0 AS [GL Account],
		a.DES_0 AS [GL Account Description],
		a.RPTCODDEB_0 AS [Reporting Code],
		b.CCE3_0 AS [Location Dim 3],
		b.AMT_0 AS [1],
		b.AMT_1 AS [2],
		b.AMT_2 AS [3],
		b.AMT_3 AS [4],
		b.AMT_4 AS [5],
		b.AMT_5 AS [6],
		b.AMT_6 AS [7],
		b.AMT_7 AS [8],
		b.AMT_8 AS [9],
		b.AMT_9 AS [10],
		b.AMT_10 AS [11],
		b.AMT_11 AS [12],
		b.AMT_12 AS [13],
		g.[CONTROLLABLE PROFIT] AS [CONTROLLABLE PROFIT],
		g.[CONTROLLABLE PROFIT - IHOP] AS [CONTROLLABLE PROFIT - IHOP],
		g.[CONTROLLABLE PROFIT - Papa John's] AS [CONTROLLABLE PROFIT - Papa John's],
		g.[NET EBITDA - Landlord] AS [NET EBITDA - Landlord],
		g.[NET EBITDA - Landlord PJ] AS [NET EBITDA - Landlord PJ],
		g.[NET PROFIT (LOSS)] AS [NET PROFIT (LOSS)],
		g.[NET PROFIT (LOSS) - FRE] AS [NET PROFIT (LOSS) - FRE],
		g.[NET PROFIT (LOSS) - IHOP] AS [NET PROFIT (LOSS) - IHOP],
		g.[NET PROFIT (LOSS) - Papa John's] AS [NET PROFIT (LOSS) - Papa John's],
		g.[NET PROFIT (LOSS) - Detail] AS [NET PROFIT (LOSS) - Detail],
		g.[NET PROFIT (LOSS) - Detail FRE] AS [NET PROFIT (LOSS) - Detail FRE],
		g.[NET PROFIT (LOSS) - Detail Papa John's] AS [NET PROFIT (LOSS) - Detail Papa John's]
	FROM PROD.BUD b
	LEFT OUTER JOIN PROD.COMPANY AS c
		ON b.CPY_0 = c.CPY_0
	LEFT JOIN PROD.GACCOUNT AS a
		ON b.ACC_0 = a.ACC_0
	LEFT OUTER JOIN dbo.[GL Accounts] AS g
		ON b.ACC_0 = g.[GL Account]
	WHERE b.FIY_0 >'18'
	) d
	UNPIVOT
		(
		[Ledger Posting Amount] FOR [Fiscal Period] IN
		([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13])
		) AS unpvt;
GO


