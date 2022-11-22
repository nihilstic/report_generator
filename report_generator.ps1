﻿# Table background color : RGB 0:94:184
## Infos ##
$app_name="APP"
$date=Get-Date -Format "MMyy"

#-----------------------------CODE-------------------------------#

## WORD INIT ##
$folder_path="$env:USERPROFILE\Desktop\Rapport\"
$word = New-Object -comobject Word.Application;$word.Visible = $false
$template="$folder_path\template.docx"
$doc = $word.Documents.Open(“$folder_path\template.docx”)
$sel=$word.selection

## Functions ##
function FindRange($search)
{
    $paras = $doc.Paragraphs
    foreach ($para in $paras){
        if ($para.Range.Text -match $search){
            $startPosition = $para.Range.Start
            $endPosition = $para.Range.Start
            return $startPosition
            return $endPosition
        }
    }
}

## General ##

$sel.Find.Execute("<APP>",$false,$true,$false,$false,$false,$true,1,$false,$app_name,1)
$sel.Find.Execute("<DATE>",$false,$true,$false,$false,$false,$true,1,$false,$date,1)
$sel.Find.Execute("<RESUME_TABLE>",$false,$true,$false,$false,$false,$true,1,$false,$null,1)
$sel.Find.Execute("<RESUME_TABLE>",$false,$true,$false,$false,$false,$true,1,$false,$null,1)

## Data Imports ##
$all_resume=Import-Csv "$folder_path\all_resume.csv" -Delimiter ";" -Encoding UTF7
$resume_table = $all_resume | Out-GridView -OutputMode Multiple 
## Resume table ##
$resume_rows=($resume_table.count+1)
$resume_columns=($resume_table|gm -MemberType NoteProperty).count
FindRange("<RESUME_TABLE>")
$range = $doc.Range($startPosition, $endPosition)
$table = $doc.Tables.Add($range,$resume_rows,$resume_columns) 
$table.Style = "resume_table"
$table.cell(1,1).range.text = "Index"
$table.cell(1,2).range.text = "Vulnérabilité"
$table.cell(1,3).range.text = "Risque"
$table.cell(1,4).range.text = "Impact sur les données"

for ($i=0; $i -lt $resume_table.Count; $i++){
    $table.cell(($i+2),1).Range.Text = $resume_table[$i].Index
    $table.cell(($i+2),2).Range.Text = $resume_table[$i].Vulnérabilité
    $table.cell(($i+2),3).Range.Text = $resume_table[$i].Risque
    $table.cell(($i+2),4).Range.Text = $resume_table[$i].'Impact sur les données'
}
$sel.Find.Execute("<RESUME_TABLE>",$false,$true,$false,$false,$false,$true,1,$false,$null,1)

## Vulns tables ##
$all_vulns=Import-Csv "$folder_path\all_vulns.csv" -Delimiter ";" -Encoding UTF7
$vulns_table = $all_vulns | Out-GridView -OutputMode Multiple 
FindRange("<VULN_TABLE>")
$range = $doc.Range($startPosition, $endPosition)
$table = $doc.Tables.Add($range,5,2) 
$table.Style = "vuln_table"
$table.cell(1,1).range.text = "Index $vuln" ##vuln
$table.cell(2,1).range.text = "Vulnérabilité"
$table.cell(3,1).range.text = "Niveau de risque"
$table.cell(4,1).range.text = "Risque"
$table.cell(5,1).range.text = "Fichier concernés"               

for ($i=0; $i -lt $vuln_table.Count; $i++){
    $table.cell(($i+2),2).Range.Text = $vuln_table[$i].Index
    $table.cell(($i+2),2).Range.Text = $vuln_table[$i].""
    $table.cell(($i+2),2).Range.Text = $vuln_table[$i].""
    $table.cell(($i+2),2).Range.Text = $vuln_table[$i].""
}
$sel.Find.Execute("<VULN_TABLE>",$false,$true,$false,$false,$false,$true,1,$false,$null,1)


## Export ##
$report_path="$folder_path\test.docx"
$doc.SaveAs("$report_path");$doc.Close();$word.Quit()


## Cleaning ##
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($range) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($table) | Out-Null
Remove-Variable doc,word,range,table
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
