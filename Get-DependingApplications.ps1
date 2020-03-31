<#
    GetApplicationsWithDependency.ps1
    **********************************************************************************************************

       Purpose: Find all Applications that have a certain dependency set (e.g. Java, .Net, c++ runtime)
       Frank Maxwitat, Mar 2020

    ***********************************************************************************************************
#>

#Enter the Application that is set as dependency
$Dependency = '%Java%'

#Enter your DB and DB server
$SCCMDBName = 'CM_p01'
$SCCMDBServerName = 'svrcmp01'

$objConnection = New-Object -comobject ADODB.Connection
$objRecordset = New-Object -comobject ADODB.Recordset
$con = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Initial Catalog=$SCCMDBName;Data Source=$SCCMDBServerName"
$strSQL = @"
select rel.FromApplicationCIID from v_CIAppDependenceRelations as rel
inner join fn_ListLatestApplicationCIs(1033) AS app on app.CI_ID = rel.ToApplicationCIID
where app.DisplayName like '$Dependency'
"@	

$objConnection.Open($con)
$objConnection.CommandTimeout = 0
# *********** Check If connection is open *******************
If($objConnection.state -eq 0)
{
	Write-Host "Error: Connection to database failed. "	
	Exit 1
}
else
{
    $CIID_Array = @()
	$objRecordset.Open($strSQL,$objConnection)
	$objRecordset.MoveFirst()
	$rows=$objRecordset.RecordCount
	do 
	{			
		$objRecordset.MoveNext()
		$CIID = $objRecordset.Fields.Item(0).Value
        $CIID_Array += $CIID        
	}     
	until ($objRecordset.EOF -eq $TRUE)
    $objRecordset.Close()

    foreach ($ID in $CIID_Array)
    {
$strSQL2 = @"
select DisplayName from fn_ListLatestApplicationCIs(1033) where CI_ID like '$ID'
"@	        
        $objRecordset.Open($strSQL2,$objConnection)
        If($objConnection.state -eq 0)
        {
	        Write-Host "Error: Connection to database failed. "	
	        Exit 1
        }
        else
        {
            $value = $objRecordset.Fields.Item(0).Value
            if($value)
            {
                Write-Host $value
            }
            $objRecordset.Close()
        }
    }    
}
