USE [AgioNet v1.3]
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		CarlosB
-- Create date: 2013.05.27
-- Description:	SP dbo.sys_getMailErrorList // Obtiene un listado de todas las ordenes con error de envío de correo
-- =============================================
CREATE PROCEDURE dbo.sys_getMailErrorList 	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	SELECT mailID, mailTo, mailFrom, mailSubject, mailBody
	FROM dbo.sys_sendEmail WHERE mailStatus = 'FAIL'
END
GO
