 
  Function ParseSingleReport {
    Param(
      [Parameter(Mandatory=$true)][string]$File
    )
    $cHeaders = @('Stale objects', 'Privileged Accounts', 'Trusts', 'Anomalies')
    $mHeaders = @('Inactive User Or Computer', 'Network topography', 'Object Config', 'Obsolete OS', 'Old authentication protocols', 'Provisioning', 'Replication', 'Vulnerability management', 'Account take over', 'ACL Check', 'Admin control', 'Control Path', 'Delegation Check', 'Irreversible change', 'Privilege control', 'RODC', 'Old trust protocol', 'SID Filtering', 'SIDHistory', 'Trust impermeability', 'Trust inactive', 'Trust Azure', 'Audit', 'Backup', 'Certificate take over', 'Golden ticket', 'Local group vulnerability', 'Network sniffing', 'Pass the credential', 'Password retrieval', 'Reconnaissance', 'Temporary Admins', 'Weak password')
    $fname = Split-Path $File -leaf
    [xml]$source = Get-Content -Raw $File
    $source.HealthcheckData.RiskRules.HealthcheckRiskRule | %{
      $risk = $_
      $CategoryPrettyPrint = $cHeaders | ?{ $_.Replace(' ', '').ToLower() -eq $risk.Category.ToLower() }
      $ModelPrettyPrint = $mHeaders | ?{ $_.Replace(' ','').ToLower() -eq $risk.Model.ToLower() }
      [PSCustomObject]@{
          Category = $CategoryPrettyPrint
          Model = $ModelPrettyPrint
          RuleID = $risk.RiskId.Replace('_', '-')
          Risk = $risk.Rationale
          Score = $risk.Points
      }
    }
  }
  