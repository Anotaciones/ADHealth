Invoke-WebRequest https://raw.githubusercontent.com/vletoux/pingcastle/master/Healthcheck/Rules/RuleDescription.resx -OutFile .\a
[xml]$resx = Get-Content -Encoding UTF8 -Path .\a
Remove-Item -Path .\a
$types = @('Title', 'Description', 'Solution', 'TechnicalExplanation')
$rules_data = @{}
$resx.root.data | %{
    $node = $_
    $nparts = $node.Name.Split('_')
    $ntype = $nparts[-1]
    $nid = ($nparts[0..($nparts.Length-2)] -Join '_').Replace('___', '$$$').Replace('_', '-')
    If ($types -contains $ntype ) {
        If ($nid -in $rules_data.Keys) {
            $rules_data[$nid][$ntype] = $node.Value
        } Else {
            $rules_data[$nid] = @{}
            $rules_data[$nid][$ntype] = $node.Value
        }
    }
}

$github_listing = Invoke-WebRequest -Uri https://github.com/vletoux/pingcastle/tree/master/Healthcheck/Rules
$source = $github_listing.Content
$html = New-Object -ComObject "HTMLFile";
$html.IHTMLDocument2_write($source);
$hcRules_hrefs = $html.all.tags("a") | ?{ ($_.href -like 'about:/vletoux/pingcastle/blob/master/Healthcheck/Rules/*.cs')} | %{ $_.href }
$rules_models = $hcRules_hrefs | %{
    $href = $_
    $rule_github_link = $href.Replace('about:', 'https://raw.githubusercontent.com').Replace('blob/master', 'master')
    $github_rule = Invoke-WebRequest -Uri $rule_github_link
    $rule_code = $github_rule.Content 
    $rulemodel_match = [regex]::match($rule_code,'\[RuleModel\((.*)\)\]').Groups[1].Value
    $rulemodel = $rulemodel_match.Replace('"', '').Replace(' RiskRuleCategory.', '').Replace(' RiskModelCategory.', '').Split(',')
    If ($rulemodel[0] -ne '') {
        [PSCustomObject]@{
            Category = $rulemodel[1]
            Model = $rulemodel[2]
            RuleID = $rulemodel[0].Replace('_', '-')
        }
    }
}

$backup_date = Get-Date -Format "dd-MM-yyyy HH:MM:ss"

$rdesc = $rules_models | %{ 
    $model = $_ 
    $details = $rules_data[$model.RuleID]
    [PSCustomObject]@{
        Category = $model.Category
        Model = $model.Model
        RuleID = $model.RuleID
        Title = $details['Title']
        Description = $details['Description']
        Explanation = $details['TechnicalExplanation']
        Solution = $details['Solution']
    }
} 

Rename-Item ".\Resources\RulesDescription.xml" "RulesDescriptionBackup.xml"
$rdesc | Export-Clixml -Path .\Resources\RulesDescription.xml