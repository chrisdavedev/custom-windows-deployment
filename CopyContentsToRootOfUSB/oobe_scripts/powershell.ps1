Write-Host "Hello from Powershell.ps1"

$targets = [System.EnvironmentVariableTarget]::Machine, [System.EnvironmentVariableTarget]::User, [System.EnvironmentVariableTarget]::Process
foreach ($target in $targets) {
  $envs = [System.Environment]::GetEnvironmentVariables($target)
  foreach ($entry in $envs.GetEnumerator()) {
    $k = $entry.Key
    $v = $entry.Value
    Write-Host ("{0}={1}  (Scope: {2})" -f $k, $v, $target)
  }
}

Read-Host -Prompt "Press Enter to continue"