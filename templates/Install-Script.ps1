##*===============================================
##* FUNCTIONS
##*===============================================
function write_log {
  param ([string]$log_string)
    
  $timestamp = Get-Date
  $formatted_log = "$timestamp $log_string"
    
  try {   
    Add-content $log_file -value $formatted_log -ErrorAction SilentlyContinue
  }
  catch {
    $formatted_log = "$timestamp Invalid data encountered"
    Add-content $log_file -value $formatted_log
  }
  Write-Host $formatted_log
}

function ensure_logging_path {
  if (![System.IO.File]::Exists($logging_path)) {
    New-Item -ItemType Directory -Force -Path $logging_path
  }
}

function get_user_info {
  write_log "Detecting logged-on user..."
    
  # Try primary method first
  $user_full = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty UserName
  write_log "User detected as: '$user_full'"

  try {
    $user = Split-Path -Leaf $user_full -ErrorAction SilentlyContinue
    write_log "Method 1: User identified as: $user"
  }
  catch {
    write_log "Trying alternative detection method..."
    $user = $env:username
    write_log "Method 2: User identified as: $user"
  }

  # Get user SID
  $principal = New-Object System.Security.Principal.NTAccount($user_full)
  $sid = $principal.Translate([System.Security.Principal.SecurityIdentifier]).Value
  write_log "User SID: '$sid'"

  return @{
    UserFull = $user_full
    User     = $user
    SID      = $sid
  }
}

function install_application {
  param (
    [string]$installer_path,
    [string]$app_name,
    [string]$installer_type,
    [string]$install_args
  )

  write_log "Starting $app_name installation..."

  try {
    if ($installer_type -eq "MSI") {
      if ([string]::IsNullOrEmpty($install_args)) {
        $install_args = "/qn /norestart"
      }

      $process = Start-Process "msiexec.exe" -ArgumentList "/i `"$installer_path`" $install_args" -Wait -PassThru

      switch ($process.ExitCode) {
        0 { write_log "Successfully installed $app_name" }
        1641 { write_log "Success - Restart initiated" }
        3010 { write_log "Success - Restart required" }
        default { 
          write_log "Installation failed with exit code: $($process.ExitCode)"
          throw "Installation failed"
        }
      }
    }
    else {
      if ([string]::IsNullOrEmpty($install_args)) {
        $install_args = "/silent"
      }

      $process = Start-Process $installer_path -ArgumentList $install_args -Wait -PassThru
            
      if ($process.ExitCode -eq 0) {
        write_log "Successfully installed $app_name"
      }
      else {
        write_log "Installation failed with exit code: $($process.ExitCode)"
        throw "Installation failed"
      }
    }
  }
  catch {
    $error_message = $_.Exception.Message
    write_log "$app_name installation error: $error_message"
    throw
  }
}

##*===============================================
##* VARIABLES
##*===============================================
$company = "Nexus"
$app_title = "<APP_TITLE>" # To be replaced during package creation
$version = "<VERSION>" # To be replaced during package creation
$installer_type = "<INSTALLER_TYPE>" # To be replaced during package creation
$install_args = "<INSTALL_ARGS>" # To be replaced during package creation
$script_name = (Get-Item $PSCommandPath).Basename
$script_full_name = (Get-Item $PSCommandPath).Name
$logging_path = "C:\ProgramData\$company\$app_title"
$log_file = "$logging_path\$script_name.log"
$installer_path = "$logging_path\$script_name.$($installer_type.ToLower())"

##*===============================================
##* MAIN EXECUTION
##*===============================================
write_log "==================== Script $script_full_name Starting ===================="

# Ensure logging directory exists
ensure_logging_path

# Gather system and user information
write_log "==================== Information Gathering Starting ===================="
$computer_name = $env:COMPUTERNAME
$computer_info = Get-ComputerInfo
$user_info = get_user_info
write_log "Installer Path: $installer_path"
write_log "App Name: $app_title"
write_log "Installer Type: $installer_type"
write_log "Install Args: $install_args"
write_log "Installer Version: $version"
write_log "Computer Name: $($computer_name)"
write_log "Computer Info: $($computer_info | Out-String)"
write_log "User Info: $($user_info.UserFull) $($user_info.User) $($user_info.SID)"
write_log "==================== Information Gathering Finished ===================="

# Perform installation
write_log "==================== Installation Starting ===================="
try {
  install_application -installer_path $installer_path -app_name $app_title -installer_type $installer_type -install_args $install_args
  write_log "==================== Installation Completed Successfully ===================="
}
catch {
  write_log "==================== Installation Failed ===================="
  exit 1
}

write_log "==================== Script Finished ===================="
