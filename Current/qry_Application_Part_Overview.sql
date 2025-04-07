SELECT tbl_Application_Part.ID, tbl_Application_Part.Name, tbl_Application_Part.Description, tbl_Application_Part.Path, tbl_Application_Part.Type_Programm_FK, tbl_Application_Part.Type_Functional_FK, tbl_Application_Part.Application_FK, tbl_Application_Part.Active
FROM tbl_Application_Part
WHERE (((tbl_Application_Part.Active)=True))
ORDER BY tbl_Application_Part.ID;

