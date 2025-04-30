SELECT tbl_Class_Property.ID, tbl_Class_Property.Name, tbl_Class_Property.Active
FROM tbl_Class_Property
WHERE (((tbl_Class_Property.Class_FK)=[Formulare]![110_frmClassBuilder]![cmbAddProperty_Class_FK]))
ORDER BY tbl_Class_Property.Name;

