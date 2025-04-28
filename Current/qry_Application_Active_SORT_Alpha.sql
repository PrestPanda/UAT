SELECT tbl_Application.ID AS Ausdr1, tbl_Application.Name AS Ausdr2
FROM tbl_Application
WHERE ((([tbl_Application].[Active])=True))
ORDER BY tbl_Application.Name;

