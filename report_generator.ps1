# Table background color : RGB 0:94:184
## Infos ##
$app_name="APP"
$date=Get-Date -Format "MMyy"
$redacteur = "Matthieu FOUET"

#-----------------------------CODE-------------------------------#

## WORD INIT ##
$folder_path="$env:USERPROFILE\Desktop\Rapport\"
$word = New-Object -comobject Word.Application;$word.Visible = $false
$doc = $word.Documents.Open(“$folder_path\template.docx”)
$sel=$word.selection

## Datas from Cx export ##
$crit_value = 10
$maj_value = 5
$mod_value = 3
$min_value = 2
$bp_value  = 1
$weak_point_1 = ""
$weak_point_2 = ""
$weak_point_3 = ""

## Functions ##
function FindRange($search)
{
    $paras = $doc.Paragraphs
    foreach ($para in $paras){
        if ($para.Range.Text -match $search){
            $start_position = $para.Range.start
            $end_position = $para.Range.End
            $position = @($start_position,$end_position)
            return $position
        }
    }
}

## General ##
$sel.Find.Execute("<APP>",$false,$true,$false,$false,$false,$true,1,$false,$app_name,1)
$sel.Find.Execute("<DATE>",$false,$true,$false,$false,$false,$true,1,$false,$date,1)
$sel.Find.Execute("<REDACTEUR>",$false,$true,$false,$false,$false,$true,1,$false,$redacteur,1)

## Resume Chart ##
$vuln_chart = $doc.InlineShapes(5).Chart
$crit_value = $vuln_chart.ChartData.Workbook.ActiveSheet.Rows[2].Cells[2].Formula = $crit_value
$maj_value = $vuln_chart.ChartData.Workbook.ActiveSheet.Rows[3].Cells[2].Formula = $maj_value
$mod_value = $vuln_chart.ChartData.Workbook.ActiveSheet.Rows[4].Cells[2].Formula = $mod_value
$min_value = $vuln_chart.ChartData.Workbook.ActiveSheet.Rows[5].Cells[2].Formula = $min_value
$bp_value = $vuln_chart.ChartData.Workbook.ActiveSheet.Rows[6].Cells[2].Formula = $bp_value

## Resume table ##
$all_vulns=Import-Csv "$folder_path\all_vulns.csv" -Delimiter ";" -Encoding UTF7
$resume_table = $all_vulns | Out-GridView -OutputMode Multiple -Title "Tableau résumé : Sélectionner les vulnérabilités"
$position = FindRange("<RESUME_TABLE>")
$range = $doc.Range($position[1], $position[1])
$table = $doc.Tables.Add($range,$($resume_table.count+1),4)
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
$vuln_table = $all_vulns | Out-GridView -OutputMode Multiple -Title "Tableaux détaillés : Sélectionner les vulnérabilités"
$position = $(FindRange("<VULN_TABLE>"))
$range = $doc.Range($position[1], $position[1])

for ($i=0; $i -lt $vuln_table.Count; $i++){
    $table = $doc.Tables.Add($range,5,2) 
    $table.Style = "vuln_table"
    $cellw = $table.Cell(1,1).Width;$table.Cell(1,1).Merge($table.cell(1,2));$table.Cell(1,1).Width = $($cellw * 2) 
    $table.Cell(1,1).Range.Text = "Index $($vuln_table[$i]."Index")"
    $table.Cell(2,1).Range.Text = "Vulnérabilité"
    $table.Cell(3,1).Range.Text = "Niveau de risque"
    $table.Cell(4,1).Range.Text = "Impact sur les données"
    $table.Cell(5,1).Range.Text = "Cx Url"  
    $table.Cell(6,1).Range.Text = "Description"  
    $table.Cell(2,2).Range.Text = $vuln_table[$i]."Vulnérabilité"
    $table.Cell(3,2).Range.Text = $vuln_table[$i]."Niveau de risque"
    $table.Cell(4,2).Range.Text = $vuln_table[$i]."Impact sur les données"
    $table.Cell(5,2).Range.Text = $vuln_table[$i]."Description"
    $range=$doc.Range($($table.Range.End + 1), $($table.Range.End + 1))
}
$sel.Find.Execute("<VULN_TABLE>",$false,$true,$false,$false,$false,$true,1,$false,$null,1)

## Additionals points ##
$position = $(FindRange("<ADDITIONNALS>"))
$range = $doc.Range($position[1], $position[1])


$sel.Find.Execute("<ADDITIONNALS>",$false,$true,$false,$false,$false,$true,1,$false,$null,1)

## Export ##
$report_path="$folder_path\test.docx"
$doc.SaveAs("$report_path");$doc.Close();$word.Quit()

## Cleaning ##
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($range) | Out-Null
#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($table) | Out-Null
Remove-Variable doc,word,range,table
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
