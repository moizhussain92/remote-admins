#Resolves each computer to FQDN. DNS suffix must exist in the local machine.
Function ResolveDNS ($Computers) {
    Write-Verbose -Verbose "Resolving the Computer Name..." 
    $ResolvedComputers = $Computers | foreach { 
        Try { 
            Resolve-DnsName $_ -ErrorAction Stop | select Name -ExpandProperty Name -First 1
        } 
        Catch { Write-Warning ($_.Exception.Message + ". " + "Try using FQDN.") }
    }
    if ($ResolvedComputers) {
        return $ResolvedComputers
    }
    else {
        Break
    }
}

#Pulls the items from the excel sheet starting from column 1, row 2 
Function PullList ($filepath) {

    $sheetName = "Sheet1" 
    $objExcel = New-Object -ComObject Excel.Application
    $workbook = $objExcel.Workbooks.Open($filepath)
    $sheet = $workbook.Worksheets.Item($sheetName)
    $rowMax = ($sheet.UsedRange.Rows).count
    $rowName, $colName = 1, 1
    $List = @()
    for ($i = 1; $i -le $rowMax - 1; $i++) {
        $name = $sheet.Cells.Item($rowName + $i, $colName).text
        $List += $name.trim()
    }       
    $objExcel.quit()
    return $List
   
} 

#Validates if the input is individual Computer Name or Excel sheet.
Function ValidateComputerName { 
    [CmdletBinding()]
    Param(
        
        [Parameter(Mandatory = $True,
            Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ComputerName        
    ) 
    try {
        if ($ComputerName -match "^*.xlsx$") {                       
            $pathExists = Test-Path $ComputerName
            if ($pathExists -eq $true) {
                Write-Verbose -Verbose "Pulling the Computer List..."                        
                $newComputerName = PullList ($ComputerName)                                
            }
            else { Write-Warning "Check the file path $ComputerName"}
        }
        else {
            $newComputerName = $ComputerName
        } 
        return $newComputerName
    }
    catch {
        Write-Warning $_.Exception.Message
    }
    
}