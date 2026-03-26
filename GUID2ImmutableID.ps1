<#
.SYNOPSIS
    Converts between Entra ID GUIDs and Immutable IDs (Base64 encoded format).

.DESCRIPTION
    This script provides a utility to convert between Microsoft Entra ID GUIDs and their
    Immutable ID representation (Base64 encoded format). The Immutable ID is used in Entra ID
    hybrid scenarios for maintaining object consistency across on-premises Active Directory
    and Entra ID.

    The script automatically detects the input format and performs the appropriate conversion:
    - GUID format → ImmutableID (Base64)
    - ImmutableID (Base64) → GUID format

.PARAMETER ValueToConvert
    The value to convert. Can be either:
    - A GUID in standard format (e.g., '748b2d72-706b-42f8-8b25-82fd8733860f')
    - An ImmutableID in Base64 format (e.g., 'ci2LdGtw+EKLJYL9hzOGDw==')

.EXAMPLE
    .\GUID2ImmutableID.ps1 '748b2d72-706b-42f8-8b25-82fd8733860f'
    
    Description
    -----------
    Converts the provided GUID to its Immutable ID representation.
    Output: ci2LdGtw+EKLJYL9hzOGDw==

.EXAMPLE
    .\GUID2ImmutableID.ps1 'ci2LdGtw+EKLJYL9hzOGDw=='
    
    Description
    -----------
    Converts the provided Immutable ID back to GUID format.
    Output: 748b2d72-706b-42f8-8b25-82fd8733860f

.NOTES
    Author: Steve Halligan
    Version: 1.0
    
    This is sample code provided by Microsoft. For production environments, ensure proper
    testing and validation before implementation.

    Requires: PowerShell 2.0 or later
    
    The Immutable ID is particularly useful in hybrid Entra ID scenarios where you need
    to maintain consistent object identity between on-premises AD and Entra ID.

.LINK
    https://docs.microsoft.com/en-us/azure/active-directory/hybrid/how-to-connect-install-existing-database

.RELATED
    Convert-OctetStringToGUID
    Convert-GuidToOctetString

.COPYRIGHT
    Copyright (c) 2012 Microsoft Corporation. All rights reserved.
    
    This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.
    THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED AS IS WITHOUT WARRANTY OF ANY KIND.
#>

#------------------------------------------------------------------------------   
#   
# Copyright © 2012 Microsoft Corporation.  All rights reserved.   
#   
# This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.
# THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,  
# INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.   
# We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code  
# form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software product in  
# which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the Sample Code is  
# embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits,  
# including attorneys' fees, that arise or result from the use or distribution of the Sample Code. 
#   
#------------------------------------------------------------------------------   
param([string]$valuetoconvert) 
 
function isGUID ($data) { 
    try { 
        $guid = [GUID]$data 
        return 1 
    } catch { 
        return 0 
    } 
} 

function isBase64 ($data) { 
    try { 
        $decodedII = [system.convert]::frombase64string($data) 
        return 1 
    } catch { 
        return 0 
    } 
} 

function displayhelp { 
    write-host "Please Supply the value you want converted" 
    write-host "Examples:" 
    write-host "To convert a GUID to an Immutable ID: GUID2ImmutableID.ps1 '748b2d72-706b-42f8-8b25-82fd8733860f'" 
    write-host "To convert an ImmutableID to a GUID: GUID2ImmutableID.ps1 'ci2LdGtw+EKLJYL9hzOGDw=='" 
} 
 
if ($valuetoconvert -eq $NULL) { 
    DisplayHelp 
    return 
} 

if (isGUID($valuetoconvert)) { 
    $guid = [GUID]$valuetoconvert 
    $bytearray = $guid.tobytearray() 
    $immutableID = [system.convert]::ToBase64String($bytearray) 
    write-host "ImmutableID" 
    write-host "-----------" 
    $immutableID 
} elseif (isBase64($valuetoconvert)) { 
    $decodedII = [system.convert]::frombase64string($valuetoconvert) 
    if (isGUID($decodedII)) { 
        $decode = [GUID]$decodedii 
        $decode 
    } else { 
        Write-Host "Value provided not in GUID or ImmutableID format." 
        DisplayHelp 
    } 
} else { 
    Write-Host "Value provided not in GUID or ImmutableID format." 
    DisplayHelp 
}


