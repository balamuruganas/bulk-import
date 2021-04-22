$database = "master"
$Language ="en"
$ParentFolderItem = Get-Item -Path (@{$true="$($database):\sitecore\content\Sitecore\Storefront\Home\Company"; $false="$($database):\content"}[(Test-Path -Path "$($database):\content\home")])
$TemplateItem = Get-Item "$($database):\sitecore\templates\Project\Sitecore\Content Page"

$props = @{
    Parameters = @(
        @{Name="ParentFolderItem"; Title="Choose the report root"; Tooltip="Only items from this root will be returned."; }
        @{ Name = "TemplateItem"; Title="Base Template"; Tooltip="Select the item to use as a base template for the report"; Root="/sitecore/templates/"}
    )
    Title = "Choose Root & Template"
    Description = "Choose the criteria for import."
    Width = 550
    Height = 300
    ShowHints = $true
    OkButtonName = "Proceed"
    CancelButtonName = "Abort"
    Icon = [regex]::Replace($PSScript.Appearance.Icon, "Office", "OfficeWhite", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

$result = Read-Variable @props

if($result -eq "cancel") {
    exit
}

function GetOrCreate-ItemWithTemplate
{
    Param([Sitecore.Data.Items.Item] $ParentItem,[Sitecore.Data.Items.Item] $Template, [System.String] $Language, [System.String] $Name, [bool] $CheckForName = $False, [bool] $CreateNew = $False)
	Process
    {
        if($CheckForName -eq $False)
        {
		    $ItemRequired = Get-ChildItem $ParentItem.ID -Language $Language | Where-Object { $_.TemplateId -eq $Template.ID} | Select -First 1
        }
        else
        {
            $ItemRequired = Get-ChildItem $ParentItem.ID -Language $Language | Where-Object { $_.TemplateId -eq $Template.ID -and $_.Name.Trim() -eq $Name.Trim()} | Select -First 1
        }
        if($CreateNew -eq $True)
		{
			$ItemRequired = New-Item -Path $ParentItem.Paths.Path -ItemType $Template.ID -Name $Name -Language $Language
		}
    }
    End
    {
        return $ItemRequired
    }
}

function Get-FieldValues
{
    Param([Sitecore.Data.Items.Item] $Item, [System.String] $FieldName, [System.String] $ValueToMatch)
    Process
    {
        $output =""
        
        $field = $item.Fields[$fieldName]
        $fieldTypes = $("multilist","droplist","treelist")
        if(!($fieldTypes -contains $field.TypeKey))
        {
            $output = $ValueToMatch;
        }
        else
        {
            $query = $field.Source;
            if($query.StartsWith("query"))
            {
                if($query.TrimStart("query:").Equals("./*"))
                {
                    $list = $item.Axes.SelectItems($query.TrimStart("query:"))
                }
                else
                {
                    $list = $item.Axes.SelectItems("$($item.Paths.Path)" +"/" + $query.TrimStart("query:"))
                }
            }
            $results = @()
            $values = $ValueToMatch.Split("{|}")
            $values |  ForEach-Object { 
                $this = $_;
                $MatchingItem = $list | Where-Object { $_.Name -match $this.Trim() } | Select -First 1
                if($MatchingItem)
                {            
                    $results += $MatchingItem.ID.ToString()
                }
            }
            $output = $($results -join "|");
        }
    }
    End
    {
        return $output
    }
}
function Get-ItemByName{
    Param([System.String] $ItemName)
	
	Process
    {
		$RequiredItem = GetOrCreate-ItemWithTemplate -ParentItem $ParentFolderItem -Template $TemplateItem -Name $ItemName -Language $Language -CheckForName $True
	}
	End
	{
		return $RequiredItem;
	}
}



#Upload the file on the Server in temporary folder
#It will create the folder if it is not found
$dataFolder = [Sitecore.Configuration.Settings]::DataFolder
$tempFolder = $dataFolder + "\temp\upload"
$filePath = Receive-File -Path $tempFolder -overwrite

if($filePath -eq "cancel"){
    exit
}

$resultSet =  Import-Csv $filePath

$rowsCount = ( $resultSet | Measure-Object ).Count;

    if($rowsCount -le 0){
        Remove-Item $filePath
        exit
    }
    
    Write-Log "Bulk Update Started!";
    
    foreach ( $row in $resultSet ) {
        $item = Get-ItemByName -ItemName $row.Name -ErrorAction SilentlyContinue
    
        if ($item){
         $fields = $item | Get-ItemField
         foreach($field in $fields){
             if($row -match $field) {
                 $value = Get-FieldValues -Item $item -FieldName $field -ValueToMatch $row.$field
                 $item.$field = $value
             }
         }
        }
        else {
            $logThis =  "Couldn't find: " + $row.ItemPath + " with Language Version: " + $row.Language 
            $logThis
            Write-Log $logThis
        }
    }
    
    $logInfo = "Bulk Update is Completed!";
    $logInfo
    Write-Log $logInfo

Remove-Item $filePath
