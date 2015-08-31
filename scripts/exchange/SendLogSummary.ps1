############################################################################
# Modify the following block of variables to suit your needs.
############################################################################
#
# Email information.
$From       = "Exchange System <it-exmaint@jmu.edu>"
$To         = "wrightst@jmu.edu"
$SmtpServer = "mailgw.jmu.edu"

# The date range for events.
$Start = (Get-Date).AddHours(-24)
$End = (Get-Date)

# An array of event logs from which to pull events.
$EventLogs = "Application", "System"

# An array of numeric log levels you want to retrieve.
# Log Levels:
# 1 = Critical
# 2 = Error
# 3 = Warning
# 4 = Information
$LogLevels = 1, 2, 3

# An array of servers from which you want to retrieve events.
$Servers = @(
        "hub1", "hub2",
        "mb1", "mb2", "mb3", "mb4",
        "et1", "et2", "et3", "et4"
        )

# An array of dedupids you want to filter out.
$IgnoredEvents = @(
        "ESE_214",
        "ESE_215",
        "ESE_2007",
        "ESE_629",
        "ESE_BACKUP_914",
        "FPSMC_Deployment_Agent_30001",
        "FPSMC_Deployment_Agent_40002",
        "FPSMC_Deployment_Agent_7001",
        "FSEAgent_8056",
        "GetEngineFiles_2109",
        "GetEngineFiles_6014",
        "IPMIDRV_1004",
        "LsaSrv_40968",
        "Microsoft-Windows-CEIP_1008",
        "Microsoft-Windows-DistributedCOM_10009",
        "Microsoft-Windows-DistributedCOM_10010",
        "Microsoft-Windows-PerfNet_2004",
        "Microsoft-Windows-TerminalServices-Printers_1111",
        "Microsoft-Windows-User_Profiles_Service_1511",
        "Microsoft-Windows-User_Profiles_Service_1530",
        "Microsoft-Windows-Winlogon_4005",
        "Microsoft-Windows-WinRM_10154",
        "MSExchange_ActiveSync_1008",
        "MSExchange_ActiveSync_1016",
        "MSExchange_ActiveSync_1107",
        "MSExchange_ActiveSync_1108",
        "MSExchange_ADAccess_2915",
        "MSExchange_Availability_4002",
        "MSExchange_Common_106",
        "MSExchange_Common_4999",
        "MSExchange_Control_Panel_39",
        "MSExchange_Extensibility_1050",
        "MSExchange_Search_Indexer_107",
        "MSExchange_MailTips_14003",
        "MSExchange_MailTips_14035",
        "MSExchange_Mid-Tier_Storage_5001",
        "MSExchange_OWA_99",
        "MSExchange_Web_Services_5",
        "MSExchange_Web_Services_6",
        "MSExchange_Web_Services_7",
        "MSExchange_Web_Services_22",
        "MSExchangeIS_Mailbox_Store_1114"
        "MSExchangeIS_Mailbox_Store_9823",
        "MSExchangeIS_Mailbox_Store_9877",
        "MSExchangeIS_Mailbox_Store_10036",
        "MSExchangeIS_8528",
        "MSExchangeIS_9646",
        "MSExchangeMailboxAssistants_10025",
        "MSExchangeRepl_2034",
        "MSExchangeRepl_2153",
        "MSExchangeRepl_2174",
        "MSExchange_Search_Indexer_118",
        "MSExchangeThrottlingClient_1002",
        "MSExchangeThrottlingClient_1003",
        "MSExchangeTransport_1035",
        "MSExchangeRepl_4110",
        "MSExchangeSA_9327",
        "MSExchangeTransport_12025",
        "Ntfs_57",
        "Schannel_36887",
        "Service Control_Manager_7024",
        "Service Control_Manager_7043",
        "Server_Administrator_1553",
        "Server_Administrator_2335",
        "Symantec_Network_Protection_400",
        "USER32_1076",
        "VSS_12289"
        )
#
# Stop editing here.
###########################################################################

Import-Module '.\EventLogSummary.psm1'

$ignored = New-Object System.Collections.ArrayList
$ignored.AddRange($ignoredEvents)
$logs = @()

foreach ($server in $Servers) {
    $logs += Get-EventLogSummary `
                -Verbose `
                -ComputerName $server `
                -LogName $EventLogs `
                -Start $Start `
                -End $End `
                -Level $LogLevels | ? {
                    $ignored.Contains($_.Dedupid) -eq $false
                }
}

,($logs | Sort -Property `
            @{Expression="Level";Descending=$false}, `
            @{Expression="Count";Descending=$true}, `
            Dedupid, MachineName) |
    Send-EventLogSummaryMailMessage -SmtpServer $SmtpServer -From $From -To $To

