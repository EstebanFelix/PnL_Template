USE [SQLEXPRESS]
GO

/****** Object:  View [dbo].[GL_Transaction]    Script Date: 6/25/2023 1:56:22 PM ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO










/*AND a.CCE_2 = '1915' --STORE NUMBER*/
CREATE VIEW [dbo].[GL_Transaction]
AS
SELECT      d.LEDTYP_0 AS [Ledger Type Code], 
			d.LED_0 AS [GL Ledger Code], 
			d.COA_0 AS [Chart of Acct Code], 
			d.CPY_0 AS [Company Code], x.FIY_0 AS [Fiscal Year], 
			x.PER_0 AS [Fiscal Period], 
			CASE WHEN x.PER_0 IN (1, 2, 3) THEN 'Q1' WHEN x.PER_0 IN (4, 5, 6) THEN 'Q2' WHEN x.PER_0 IN (7, 8, 9) THEN 'Q3' WHEN x.PER_0 IN (10, 11, 12, 13) THEN 'Q4' END AS [Fiscal Quarter],
			d.OFFACC_0 AS [Partner Code], 
			d.BPR_0 AS [Customer Code], 
            d.ACC_0 AS [GL Account],
			c.DES_0 AS [GL Account Description],
			c.RPTCODDEB_0 AS [Reporting Code],
		    d.FCYLIN_0 AS [Site Code],
			a.CCE_0 AS [Sub Account],
			a.CCE_1 AS [Category Dim 2],
			a.CCE_2 AS [Location Dim 3],
			d.ACCDAT_0 AS [Accounting Date],
			CAST(d.ACCDAT_0 AS VARCHAR(12)) AS [Accounting Date (Alpha)],
			d.NUM_0 AS [Document Number],
			d.DES_0 AS [Document Description],
			CASE WHEN d .SNS_0 = 1 THEN CASE WHEN a.AMTLED_0 IS NOT NULL THEN a.AMTLED_0 * (d .SNS_0)  ELSE d .AMTLED_0 * (d .SNS_0) END ELSE 0 END AS [Ledger Credit], 
			CASE WHEN d .SNS_0 = - 1 THEN CASE WHEN a.AMTLED_0 IS NOT NULL THEN a.AMTLED_0 * (d .SNS_0) ELSE d .AMTLED_0 * (d .SNS_0) END ELSE 0 END AS [Ledger Debit], 
			CASE WHEN a.AMTLED_0 IS NOT NULL THEN a.AMTLED_0 * (d .SNS_0) ELSE d .AMTLED_0 * (d .SNS_0) END AS [Ledger Posting Amount], 
            CASE WHEN d .SNS_0 = - 1 THEN CASE WHEN a.AMTLED_0 IS NOT NULL THEN a.AMTLED_0 * (d .SNS_0) * - 1 ELSE d .AMTLED_0 * (d .SNS_0) * - 1 END ELSE 0 END AS [Ledger Rev Credit],
			x.CAT_0 AS [Journal Category], 
            x.DESVCR_0 AS [Default Description],
			x.REF_0 AS Reference,
			x.JOU_0 AS [Journal Code],
			d.TYP_0 AS [Document Type],
			x.BPRDATVCR_0 AS [Document Date],
			x.BPRVCR_0 AS [Source Document],
			d.CREUSR_0 AS [Create User], 
            d.CREDATTIM_0 AS [Create Date],
			CASE WHEN d .SNS_0 = - 1 THEN CASE WHEN a.AMTCUR_0 IS NOT NULL THEN a.AMTCUR_0 * (d .SNS_0) ELSE d .AMTCUR_0 * (d .SNS_0) END ELSE 0 END AS [Transaction Credit], 
            CASE WHEN d .SNS_0 = 1 THEN CASE WHEN a.AMTCUR_0 IS NOT NULL THEN a.AMTCUR_0 * (d .SNS_0) ELSE d .AMTCUR_0 * (d .SNS_0) END ELSE 0 END AS [Transaction Debit],
			CASE WHEN a.AMTCUR_0 IS NOT NULL THEN a.AMTCUR_0 * (d .SNS_0) ELSE d .AMTCUR_0 * (d .SNS_0) END AS [Transaction Posting],
			d.SNS_0 AS Sign,
			x.RVS_0 AS Reversal,
			x.RVSDAT_0 AS [Reverse Date],
			x.STA_0 AS [Journal Status],
			g.[Balance Sheet] AS [Balance Sheet],
			GETDATE() AS [Refresh Date]

FROM				PROD.GACCENTRYD AS d LEFT OUTER JOIN
                         PROD.GACCENTRYA AS a ON d.TYP_0 = a.TYP_0 AND d.NUM_0 = a.NUM_0 AND d.LIN_0 = a.LIN_0 AND d.LEDTYP_0 = a.LEDTYP_0 LEFT OUTER JOIN
						 dbo.[GL Accounts] AS g ON d.ACC_0 = g.[GL Account] LEFT OUTER JOIN
						 PROD.GACCOUNT AS c ON d.ACC_0 = c.ACC_0 LEFT OUTER JOIN
                         PROD.GACCENTRY AS x ON d.TYP_0 = x.TYP_0 AND d.NUM_0 = x.NUM_0
WHERE				(d.ACC_0 BETWEEN '0000' AND '9999') AND (x.FIY_0>= '18')
GO