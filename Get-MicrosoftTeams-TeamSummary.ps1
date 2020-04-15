######################################################################
# Description: Get Public Teams list with Owner
# Author: Jean Luc Gauthier
# Date: 30/03/2020
# Requires: MicrosoftTeams Module
######################################################################

# Define parameters
# $Teams - Get the teams list from the tenant
# $TeamsOwner - Get teams owner
# $Output - Create a temporary PSObject to gather team details
# $TeamsReport - Output report file

# Default status message colors
$ErrMsgColor = @{BackgroundColor="Red";ForegroundColor="White"}
$SuccesMsgColor = @{BackgroundColor="Green";ForegroundColor="Black"}

# Initialize the output report file
$TeamsReport =  @()

# Check requierements
Function ModuleCheck{
    if (Get-Module -ListAvailable -Name MicrosoftTeams) {
        Write-Host "Le module Powershell pour Teams est installé." @SuccesMsgColor
    } else {
        Write-Host "Le module Powershell pour Teams est requis, nous essayons de l'installer" @SuccesMsgColor
        Try {
            Install-Module MicrosoftTeams -Force -AllowClobber
        }
        Catch{
            Write-Host "Impossible d'installer le module Powershell Teams" @ErrMsgColor
            Break
        }
    }

    If (!(Get-Module MicrosoftTeams)) {
        Write-Host "Nous chargeons le module Powershell Teams" @SuccesMsgColor
        try {
            Import-Module MicrosoftTeams    
        }
        catch {
            Write-Host "Impossible de charger le module Powershell Teams" @ErrMsgColor
            Break
        }
    
    } else {
        Write-Host "Le module Powershell Teams est déjà chargé" @SuccesMsgColor
    }
}

# Function to retrieve Public Teams Ownership details
function GetTeamsSummary {
    try {
        $Teams = Get-Team -Visibility Public
            $Teams | ForEach-Object {
            $TeamsOwner = Get-TeamUser -GroupId $_.GroupId -Role Owner

            # Add details to report object
            $Output = New-Object -TypeName PSobject
            $Output | add-member NoteProperty "Nom" -value $_.DisplayName
            $Output | add-member NoteProperty "Description" -value $_.Description
            $Output | add-member NoteProperty "Responsable" -value $TeamsOwner.Name
            $Output | add-member NoteProperty "Courriel" -value $TeamsOwner.User

            # Add oject details to report file
            $TeamsReport += $Output
        }
    }
    catch {
        Write-Host "Unable to get details for $_" @DisplayName " Error detected: $_" @ErrMsgColor
    }
    $TeamsReport
}

# Main block
ModuleCheck
GetTeamsSummary