SELECT tbl_Package_Class.*, tbl_Package.Name
FROM tbl_Package_Class INNER JOIN tbl_Package ON tbl_Package_Class.Package_FK = tbl_Package.ID
WHERE (((tbl_Package.Name)=[Forms]![Package_Management]![cmbPackageManage_Name]));

