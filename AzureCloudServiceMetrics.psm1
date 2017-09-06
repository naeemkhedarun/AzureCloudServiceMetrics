$ErrorActionPreference = "Stop"

function Export-AzureCloudServices {

    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)] 
        [string] $subscriptionName,
        [Parameter(Mandatory = $true)]
        [bool] $isProduction
    )

    Select-AzureRmSubscription -SubscriptionName $subscriptionName
    Select-AzureSubscription -SubscriptionName $subscriptionName

    $getAzureDeployments = [scriptblock] {    
        $production = Get-AzureDeployment -ServiceName $args.ResourceName -Slot Production -ErrorAction Continue
        $staging = Get-AzureDeployment -ServiceName $args.ResourceName -Slot Staging -ErrorAction Continue
        @{ Service = $production; Resource = $args }
        @{ Service = $staging; Resource = $args }
    }

    $deployments = Get-AzureRmResource `
        | ? ResourceType -eq "Microsoft.ClassicCompute/domainNames" `
        | Invoke-Parallel $getAzureDeployments `
        | ? { $_.Service -ne $null -and $_.Service.RoleInstanceList[0].InstanceSize -ne $null} `
        | % { [pscustomobject] @{
            "CloudServiceName" = $_.Service.ServiceName;
            "Count" = $_.Service.RoleInstanceList.Count;
            "Size" = $_.Service.RoleInstanceList[0].InstanceSize;
            "RoleName" = $_.Service.RoleInstanceList[0].RoleName;
            "Slot" = $_.Service.Slot;
            "ResourceId" = $_.Resource.ResourceId;
            "Region" = $_.Resource.Location;
            "Status" = $_.Service.Status;
            "HasWebRoles" = (($_.Service.RoleInstanceList[0].InstanceEndpoints | ? { $_.Protocol -iin @("http", "https") }).length -gt 0);
            "IsProduction" = $isProduction;
        }}

    $deployments | Export-Csv .\$subscriptionName.csv -NoTypeInformation
}

function Export-ComputeUtilisationForSubscription {
    param($subscriptionName)

    $deployments = Import-Csv .\$subscriptionName.csv 

    $getMetrics = {
        $date = [DateTime]::UtcNow;
        $metric = Get-AzureRmMetric -ResourceId "$($args.ResourceId)/slots/$($args.Slot)/roles/$($args.RoleName)" `
            -TimeGrain ([TimeSpan]::FromHours(1)) `
            -StartTime $date.Date.AddDays(-1) `
            -EndTime $date.Date.AddSeconds(-1) `
            -MetricNames "Percentage CPU"
        
        $utilisation = [pscustomobject] @{ ResourceId = $($args.ResourceId); Slot = $($args.Slot); RoleName = $($args.RoleName); }
        $metric.MetricValues | % { $utilisation | Add-Member @{ $_.Timestamp.Hour = $_.Average } } 

        return $utilisation;
    }

    $deployments | Invoke-Parallel $getMetrics | Export-Csv .\$subscriptionName.Metrics.csv -NoTypeInformation
}

function Merge-Subscriptions {

    $subscriptionNames = gci *.metrics.csv | % { $_.Name.Replace(".Metrics.csv", "") }

    $subscriptionNames | % { "$_.csv" } | Import-Csv | Export-Csv "Merged.csv" -NoTypeInformation
    $subscriptionNames | % { "$_.Metrics.csv" } | Import-Csv | Export-Csv "Merged.Metrics.csv" -NoTypeInformation
}


function Invoke-Parallel { 
    Param(
        [Parameter(Mandatory = $true)] [ScriptBlock]$ScriptBlock, 
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)] $ObjectList,
        $MaxThreads = 20,
        $SleepTimer = 200,
        $MaxResultTime = 120
    )
 
    Begin {
        $ISS = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
        $RunspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads, $ISS, $Host)
        $RunspacePool.Open()
        $Code = $ScriptBlock
        $Jobs = @()
    }
 
    Process {
        Write-Progress -Activity "Preloading threads" -Status "Starting Job $($jobs.count)"
        ForEach ($Object in $ObjectList) {
            $PowershellThread = [powershell]::Create().AddScript($Code)
            $PowershellThread.AddArgument($Object) | out-null
            $PowershellThread.RunspacePool = $RunspacePool
            $Handle = $PowershellThread.BeginInvoke()
            $Job = "" | Select-Object Handle, Thread, object
            $Job.Handle = $Handle
            $Job.Thread = $PowershellThread
            $Job.Object = $Object
            $Jobs += $Job
        }
        
    }
 
    End {
        $ResultTimer = Get-Date
        While (@($Jobs | Where-Object {$_.Handle -ne $Null}).count -gt 0) {
    
            $Remaining = "$($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False}).object)"
            If ($Remaining.Length -gt 60) {
                $Remaining = $Remaining.Substring(0, 60) + "..."
            }
            Write-Progress `
                -Activity "Waiting for Jobs - $($MaxThreads - $($RunspacePool.GetAvailableRunspaces())) of $MaxThreads threads running" `
                -PercentComplete (($Jobs.count - $($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False}).count)) / $Jobs.Count * 100) `
                -Status "$(@($($Jobs | Where-Object {$_.Handle.IsCompleted -eq $False})).count) remaining - $remaining" 
 
            ForEach ($Job in $($Jobs | Where-Object {$_.Handle.IsCompleted -eq $True})) {
                $Job.Thread.EndInvoke($Job.Handle)
                $Job.Thread.Dispose()
                $Job.Thread = $Null
                $Job.Handle = $Null
                $ResultTimer = Get-Date
            }
            If (($(Get-Date) - $ResultTimer).totalseconds -gt $MaxResultTime) {
                Write-Error "Child script appears to be frozen, try increasing MaxResultTime"
                Exit
            }
            Start-Sleep -Milliseconds $SleepTimer
        
        } 
        $RunspacePool.Close() | Out-Null
        $RunspacePool.Dispose() | Out-Null
    } 
}

