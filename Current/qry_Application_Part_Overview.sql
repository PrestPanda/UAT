SELECT tbl_Application_Part.ID AS Ausdr1, tbl_Application_Part.Name AS Ausdr2, tbl_Application_Part.Description AS Ausdr3, tbl_Application_Part.Path AS Ausdr4, tbl_Application_Part.Type_Programm_FK AS Ausdr5, tbl_Application_Part.Type_Functional_FK AS Ausdr6, tbl_Application_Part.Application_FK AS Ausdr7, tbl_Application_Part.Active AS Ausdr8
FROM tbl_Application_Part
WHERE ((([tbl_Application_Part].[Active])=True))
ORDER BY tbl_Application_Part.ID;

