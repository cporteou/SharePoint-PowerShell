
function Export-SPLibrary
{
	<#
		.SYNOPSIS
		Copy the contents of a SharePoint document library to disk

		.DESCRIPTION
		This function copies all files from within the chosen SharePoint libraries to a folder. Use of the -Recurse switch will also include all subfolders and their contents

		.PARAMETER Url
		URL of the SharePoint site 

		.PARAMETER Path
		Output location for exporting files and folders

		.PARAMETER Library
		This is the name of the SharePoint libraries to be exported. This should be their name and NOT their title
		
		.PARAMETER Recurse
		Switch to iterate through all subfolders of the provided library/Libraries

		.EXAMPLE
		Export-SPLibrary -Url https://SharePointURL/Sites/SiteName -Path "C:\SharePointDocuments\Output" -Library "Library1","Library2"

		.EXAMPLE
		Export-SPLibrary -Url https://SharePointURL/Sites/SiteName -Path "C:\SharePointDocuments\Output" -Library "Library1","Library2" -Recurse -Verbose

		.NOTES
			Author: Craig Porteous
			Created: July 2018
			Based on script written by Anatoly Mironov
			https://github.com/mirontoli/sp-lend-id/blob/master/aran-aran/Pull-Documents.ps1
	#>

	[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
	Param(
		[Parameter(Mandatory=$true)]
		[string]
		$Url,
		[Parameter(Mandatory=$true)]
		[ValidateScript({Test-Path -Path $_})]
		[string]
		$Path,
		[Parameter(Mandatory=$true)]
		[String[]]
		$Library,
		[switch]
		$Recurse
	)

	#! Is this needed?	
	[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")

	Write-Verbose "Connecting to SP Site to retrieve Web"
	$site = new-object microsoft.sharepoint.spsite($Url)
	$web = $site.OpenWeb()
	$site.Dispose()

	for ($i=0; $i -lt $Library.Length; $i++)
	{
		if(!$web.GetFolder($Library[$i]))
		{			
			Write-Error "The '$Library[$i]' library cannot be found"
			$web.Dispose()
			return
		}
		else 
		{
			Write-Verbose "Retrieving library:$($Library[$i]) from SharePoint"
			$folder = $web.GetFolder($Library[$i])

			# Create local path
			$rootDirectory = $Path
			$directory = Join-Path $Path $folder.Name

			if (Test-Path $directory) 
			{
				#! Put in an Overwrite option here 
				Write-Error "The folder $Library in the current directory already exists, please remove it"
				$web.Dispose()
				return
			}
			else 
			{
				$fileArray = @()

				# $fileCount = Get-SPFileCount $folder
				if($PSCmdlet.ShouldProcess($folder, "Copying documents from Library"))
				{
					if($Recurse)
					{
						Write-Verbose "Saving $fileCount files from $($folder.Name) and subfolders to $directory"
						$fileArray = Save-SPLibrary $folder $rootDirectory -Recurse 
						$fileArray | Export-Csv -Path "$($rootDirectory)\$($folder.Name).csv" -NoTypeInformation
					}
					else 
					{
						Write-Verbose "Saving $fileCount files from $($folder.Name) to $directory"
						$fileArray = Save-SPLibrary $folder $rootDirectory
						$fileArray | Export-Csv -Path "$($rootDirectory)\$($folder.Name).csv" -NoTypeInformation
					}
				}
			}			
			# $fileCount = Get-SPFileCount $folder
		}

		$web.Dispose()

	}
}

function Save-SPFile
{
	<#
		.SYNOPSIS
		Saves a specified document in SharePoint to disk

		.DESCRIPTION
		This function takes a SharePoint file object and saves a copy to disk. It is fed from the Export-SPLibrary function

		.PARAMETER File
		SharePoint File object.

		.PARAMETER Path
		Output path to save file to

		.EXAMPLE
		Save-SPFile -File $file -Path "C:\SharePointOutput\"

		.LINK
		Export-SPLibrary

		.LINK
		Save-SPLibrary
	#>

	[CmdletBinding()]
	Param(
		[Parameter(Mandatory=$true)]
		[psObject]
		$File,
		[Parameter(Mandatory=$true)]
		[string]
		$Path
	)

	$data = $File.OpenBinary()
	$Path = Join-Path $Path $File.Name
	# progress $path
	[System.IO.File]::WriteAllBytes($Path, $data) 
	Write-Verbose "$($File.Name) saved to disk: $Path"
	#! Can we update the file properties?

	return $Path
}

function Save-SPLibrary
{
	<#
		.SYNOPSIS
		Saves the contents of a specified SharePoint library to disk
		
		.DESCRIPTION
		This function takes a SharePoint library object and calls the Save-SPFile function to save the contents of the library to disk. The recurse switch will include all subfolders and files
		
		.PARAMETER Folder
		SharePoint Library(folder) object.
		
		.PARAMETER Path
		Output location for exporting files and folders
		
		.PARAMETER Recurse
		Switch to iterate through all subfolders of the provided library(folder)
		
		.EXAMPLE
		Save-SPLibrary -Folder $folder -Path "C:\SharePointOutput\" -Recurse
		
		.LINK
		Export-SPLibrary

		.LINK
		Save-SPFile
	#>
	[CmdletBinding()]
	Param(
		[Parameter(Mandatory=$true)]
		[psObject]
		$Folder,
		[Parameter(Mandatory=$true)]
		[string]
		$Path,
		[switch]
		$Recurse
	)

	#Target directory
	$directory = Join-Path $Path $Folder.Name
	#Forms folder is not wanted.
	if ($Folder.Name -eq 'Forms') {
		return
	}
	#Logging Array
	$spLog = @()
	
	if($Folder.Files.Count -gt 0 -or $Folder.SubFolders.Count -gt 0)
	{
		#Only creating directories that contain files or subfolders
		mkdir $directory | Out-Null		
		
		foreach($file in $Folder.Files)
		{	
			#Saving file to directory		
			$localName = Save-SPFile $file $directory

			Write-Verbose "Adding to log."
			$spArray = @()
			$spArray += New-Object PSObject -Property @{
				FileName="$($file.Name)"
				SPRelativeUrl="$($file.ServerRelativeUrl)"
				SPParentFolder="$($file.ParentFolder)"
				LocalFileName="$($localName)"
			}
			$spLog += $spArray
		}		
		if($Recurse)
		{
			$Folder.Subfolders | Foreach-Object { Save-SPLibrary $_ $directory -Recurse }
		}
	}
	return $spLog
}