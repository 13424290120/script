Try  
{  
    Import-module "sql" -ErrorAction Stop
    Write-Host "Succeed"  
}  
  
catch
{  
    Write-Verbose -Verbose "Failed to import the module, please verify module installation"
    Write-Verbose -Verbose $Error[0] 
}  
