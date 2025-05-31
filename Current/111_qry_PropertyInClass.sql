SELECT tbl_Property.ID, tbl_Property.Name
FROM tbl_Property
WHERE (((tbl_Property.Class_FK)=[Forms]![110_frmClassBuilder]![cmbAddProperty_Class_FK]))
ORDER BY tbl_Property.Name;

