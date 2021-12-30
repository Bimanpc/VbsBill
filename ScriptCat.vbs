Set objAdminIS = CreateObject("Microsoft.ISAdm")
objAdminIS.Stop()

Set objCatalog = objAdminIS.AddCatalog("Scripts Catalog","c:\scripts")
objAdminIS.Start