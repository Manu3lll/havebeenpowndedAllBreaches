<#
.DESCRIPTION
  Script to Output a Excel List with all Breaches from haveibeenpwned.com
.INPUTS
  Input Path to folder (without filename)
.OUTPUTS
 Excel List to given path with filename Breaches.xlsx
.NOTES
  Version:        1.0
  Author:         Manuel Hiller
  Creation Date:  16.04.2021
  E-Mail:         manuel.hiller@thak.de
#>



##############
#Excel
#############
$excel = New-Object -ComObject excel.application
$excel.visible = $False
$workbook = $excel.Workbooks.Add()
$diskSpacewksht= $workbook.Worksheets.Item(1)
############
#Add Headings
###########
$diskSpacewksht.Name = 'Breaches'
$diskSpacewksht.Cells.Item(1,1) = 'Name'
$diskSpacewksht.Cells.Item(1,2) = 'Breachdatum'
$diskSpacewksht.Cells.Item(1,3) = 'Hinzugefügt am'
$diskSpacewksht.Cells.Item(1,4) = 'Betroffene Accounts'
$diskSpacewksht.Cells.Item(1,5) = 'Gehackte Informationen'
#set Start Column
$col = 2

#Request to Haveibeenpwned Rest Api
$Url = "https://haveibeenpwned.com/api/v3/breaches/" 
$request = Invoke-RestMethod -Uri $Url -UseDefaultCredentials
    
    #For each Object in $request
    foreach ($ob in $request)
    {
    #Add Attribute Value to Excel Column
        $diskSpacewksht.Cells.Item($col,1) = $ob.name
        $diskSpacewksht.Cells.Item($col,2) = $ob.BreachDate       
        $diskSpacewksht.Cells.Item($col,3) = $ob.AddedDate     
        $diskSpacewksht.Cells.Item($col,4) = ""+$ob.PwnCount
        $string = ""
        #For each Dataclass
        foreach($i in $ob.DataClasses)

        {
        #Prevent first comma
        if($string -eq ""){
            $string = $i

        }else
        {
            $string = $string+", "+$i
        }



    }
    #Add all Dataclasses to column
    $diskSpacewksht.Cells.Item($col, 5) = $string

    #increase col         
    $col++




}

#Output path -> Filename Breaches.xlsx
$filename = "Breaches.xlsx"
[string]$path = Read-Host "Input path to folder where the file should be placed (e.g C:\temp\), please"
if(Test-path -Path $path)
{
Write-Output "Path is correct, writing file to "+$path+""+$filename
$workbook.SaveAs($path+""+$filename)  
$workbook.Close
$excel.DisplayAlerts = 'False'
$excel.Quit()
}