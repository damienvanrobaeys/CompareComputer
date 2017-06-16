$Global:Current_Folder = split-path -parent $MyInvocation.MyCommand.Definition
$Date = get-date -format "dd-MM-yy_HHmm"
$CompName = $env:COMPUTERNAME
$Vendor = (gwmi win32_computersystemproduct).vendor

New-Alias VSS  Compare-Services
New-Alias VSD  Compare-Drivers
New-Alias VSP  Compare-Process
New-Alias VSB  Compare-BIOS
New-Alias VSSo Compare-Software
New-Alias VSA  Compare-All

$Temp = $env:temp					
$Temp_CSV_Values = $Temp + "\" + "CSV_Temp_Values"						
$Temp_CSV_Equal = $Temp_CSV_Values + "\" + "Equal_Values.csv"
$Temp_CSV_MissingInFile1 = $Temp_CSV_Values + "\" +  "Values_missing_file1.csv"
$Temp_CSV_MissingInFile2 = $Temp_CSV_Values + "\" +  "Values_missing_file2.csv"
$Temp_CSV_DiffVersion = $Temp_CSV_Values + "\" + "Diff_versions.csv"
$CSS_File = "$Current_Folder\Master_Export_Compare.css" # CSS for HTML Export
$Module_version = "1.0.0.0"

If (test-path $Temp_CSV_Values)
	{
		remove-item $Temp_CSV_Values -recurse -force
		new-item $Temp_CSV_Values -type directory | out-null
	}		
Else 
	{
		new-item $Temp_CSV_Values -type directory | out-null
	}		



<#.Synopsis
	The Compare-Services function allows you to compare two computer configurations using two CSV or XML files.
	It works like this: Compare-Services -Path C:\ -file1 <File1_Path> -file2 <File2_Path> -CSV or -XLS
.DESCRIPTION
	Allow you to compare two computer configurations by comparing twi CSV or XML files.
	These CSV or XML files have to contain bith services export with Name, Caption, State and Startmode columns
	They could be generated easily using the ConfigExport module
	Comparison can be done in two formats: HTML or XLSX

.EXAMPLE
PS Root\> Compare-Services -Path C:\ -file1 c:\Services1.csv -file2 c:\Services2.csv -HTML
The command above will compare both Services1.csv and Services2.csv files. 
It will create a file Compare_Services.html file in C:

.EXAMPLE
PS Root\> Compare-Services -Path C:\ -file1 c:\Services1.csv -file2 c:\Services2.csv -XLS
The command above will compare both Services1.csv and Services2.csv files. 
It will create a file Compare_Services.xlsx file in C:

.NOTES
    Author: Damien VAN ROBAEYS - @syst_and_deploy - http://www.systanddeploy.com
#>	
	
Function Compare-Services
{
[CmdletBinding()]
Param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true, position=1)]
        [string] $Path,
        [Parameter(Mandatory=$true)]		
        [string] $File1,
        [Parameter(Mandatory=$true)]		
        [string] $File2,		
        [Switch] $XLS,
        [Switch] $HTML		
      )
    
    Begin
    {		
		# If both files File1 and File2 are CSV
		If (($File1.contains("csv")) -and ($File2.contains("csv")))
			{
				$file1_content = import-csv $File1	 
				$file2_content = import-csv $File2							
			}
			
		# If both files File1 and File2 are XML			
		ElseIf (($File1.contains("xml")) -and ($File2.contains("xml")))
			{
				$file1_content = import-clixml $File1	 
				$file2_content = import-clixml $File2							
			}
			
		# If File1 is XML and File2 is CSV			
		ElseIf (($File1.contains("xml")) -and ($File2.contains("csv")))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You have specified two different files to compare" -foregroundcolor "yellow"		
				write-host " !!! The first file is a XML and the second file is an CSV" -foregroundcolor "yellow"		 								
				write-host " !!! You need to specify two CSV or two XML files for the comparison" -foregroundcolor "yellow"										
				write-host "*******************************************************************"			
				break
			}
			
		# If File1 is CSV and File2 is XML						
		ElseIf (($File1.contains("csv")) -and ($File2.contains("xml")))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You have specified two different files to compare" -foregroundcolor "yellow"		
				write-host " !!! The first file is a CSV and the second file is an XML" -foregroundcolor "yellow"	
				write-host " !!! You need to specify two CSV or two XML files for the comparison" -foregroundcolor "yellow"										
				write-host "*******************************************************************"		
				break				
			}					
		
		# If there is no output format specified
		If ((-not $XLS) -and (-not $HTML))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You need to specific an output format. " -foregroundcolor "yellow"		
				write-host " !!! Use the switch -XLS to export in CSV or XLS format " -foregroundcolor "yellow"		 	
				write-host " !!! Use the switch -html to export in HTML format " -foregroundcolor "yellow"										
				write-host "*******************************************************************"					
			}			

		# If there output format is XLS			
		ElseIf($XLS)
			{			
				Try
					{
						# Check if an Excel process is already running, if yes it will keep the ID in the variable $excel_ID_Before
						$Excel_Process_Before = Get-Process Excel -ErrorAction SilentlyContinue
						$Excel_ID_Before = $Excel_Process_Before.id	
											
						$Excel_Test = new-object -comobject excel.application
						$Excel_value = $True 
												
						write-host ""		
						write-host "********************************************************************************************************"	
						write-host "Services will be compared in XLS format" -foregroundcolor "Cyan"							
					}

				Catch 
					{
						$Excel_value = $false 					
					}		
			}	
			
		# If there output format is HTML			
		ElseIf($HTML)
			{
				write-host ""		
				write-host "********************************************************************************************************"	
				write-host "Services will be exported in HTML format" -foregroundcolor "Cyan"		
			}				
    }
	

    Process
    {
		If($XLS)
			{
				If ($Excel_value -eq $True)
					{											
						$Compare_Services_ToXLS = $Path + "\" + "Compare_Services.xlsx"
						
						$Temp_CSV_DiffStartmode = $Temp_CSV_Values + "\" +  "Differents_startmode.csv"
						$Temp_CSV_DiffState = $Temp_CSV_Values + "\" +  "Differents_state.csv"
						
						$nbnewfile1 = 0
						$nbnewfile2 = 0		
						$nbdiffstate = 0
						$nbdiffstartmode = 0		
						$nbsame = 0	

						$equal = compare-object $file1_content $file2_content -property name, state, startmode -includeequal  | Where {$_.SideIndicator -eq "=="} | 
						   Group-Object -Property name | % { New-Object psobject -Property @{		
								Name=$_.name
								Start_Mode=$_.group[0].startmode
								State=$_.group[0].state			
							}}  | Select Name, Start_Mode, State | export-csv -encoding UTF8 -notype  $Temp_CSV_Equal 

						$nbsame = (compare-object $file1_content $file2_content -property name, state, startmode -includeequal  | Where {$_.SideIndicator -eq "=="}|measure-object).count 		
						
						### NEW SERVICES IN FILE 1
						$Result_newIn1 = "<p class=New_object>New services in $file1_name</p>"						
						$found = $false
						$newin1 = Foreach ($line1 in $file1_content)
							{
								$found = $false
								ForEach ($line2 in $file2_content)
									{
										IF ($line1.name -eq $line2.name)
											{
												$found = $true
												break
											}
									}

								IF (-not $found) 
									{
										New-Object -TypeName PSObject -Property @{
											Name = $line1.name
											State = $line1.state
											Startmode = $line1.startmode 	
											}
										$nbnewfile1 = $nbnewfile1 + 1
									}
							}	

						### NEW SERVICES IN FILE 2
						$Result_newIn2 = "<p class=New_object>New services in $file2_name</p>"						
						$found = $false
						$newin2 = Foreach ($line2 in $file2_content)
							{
								$found = $false
								ForEach ($line1 in $file1_content)
									{
										IF ($line1.name -eq $line2.name)
											{
												$found = $true
												break
											}
									}

								IF (-not $found) 
									{
										New-Object -TypeName PSObject -Property @{
											Name = $line2.name
											State = $line2.startmode
											Startmode = $line2.startmode 	
											}
										$nbnewfile2 = $nbnewfile2 + 1
									}
							}								

						### SAME SERVICES BUT DIFFERENT STATE
						$Diff_state = ForEach ($line1 in $file1_content)
						{
							ForEach ($line2 in $file2_content)
								{
									IF ($line1.name -eq $line2.name)
										{
											IF ($line1.state -ne $line2.state)
												{		
													New-Object -TypeName PSObject -Property @{
														Name = $line1.name
														State_File1 = $line1.state
														State_File2 = $line2.state 
														}  
													$nbdiffstate = $nbdiffstate + 1
												}
												Break
										}
								}												
						}


						### SAME SERVICES BUT DIFFERENT STARTMODE
						$Diff_startmode = ForEach ($line1 in $file1_content)
						{
							ForEach ($line2 in $file2_content)
								{
									IF ($line1.name -eq $line2.name)
										{
											IF ($line1.startmode -ne $line2.startmode)
												{		
													New-Object -TypeName PSObject -Property @{
														Name = $line1.name
														Startmode_F1 = $line1.startmode
														Startmode_F2 = $line2.startmode 
														}  
													$nbdiffstartmode = $nbdiffstartmode + 1		
												}
												Break
										}
								}												
						}

						$newin1 | Select Name, State, Startmode | export-csv -encoding UTF8 -notype $Temp_CSV_MissingInFile2
						$newin2 | Select Name, Startmode, Startmode | export-csv -encoding UTF8 -notype $Temp_CSV_MissingInFile1						
						$Diff_startmode | select Name, Startmode_F1, Startmode_F2 | export-csv -encoding UTF8 -notype  $Temp_CSV_DiffStartmode		
						$Diff_state | select Name, State_File1, State_File2 | export-csv -encoding UTF8 -notype  $Temp_CSV_DiffState	
																		
						$xl = new-object -comobject excel.application
						$xl.visible = $false
						$xl.DisplayAlerts=$False							
							
						$Workbook1 = $xl.workbooks.open($Temp_CSV_Equal)
						$Workbook2 = $xl.workbooks.open($Temp_CSV_MissingInFile1)
						$Workbook3 = $xl.workbooks.open($Temp_CSV_MissingInFile2) 
						$Workbook4 = $xl.workbooks.open($Temp_CSV_DiffState) 
						$Workbook5 = $xl.workbooks.open($Temp_CSV_DiffStartmode) 								

						$WorkBook0 = $xl.WorkBooks.add()

						$sh1_wborkbook0 = $WorkBook0.sheets.item(1) # first sheet in destination workbook
						$sheetToCopy1 = $Workbook1.sheets.item(1) # source sheet to copy
						$sheetToCopy1.copy($sh1_wborkbook0) # copy source sheet to destination workbook

						$sh2_wborkbook0 = $WorkBook0.sheets.item(2) # first sheet in destination workbook
						$sheetToCopy2 = $Workbook2.sheets.item(1) # source sheet to copy
						$sheetToCopy2.copy($sh2_wborkbook0) # copy source sheet to destination workbook

						$sh3_wborkbook0 = $WorkBook0.sheets.item(3) # first sheet in destination workbook
						$sheetToCopy3 = $Workbook3.sheets.item(1) # source sheet to copy
						$sheetToCopy3.copy($sh3_wborkbook0) # copy source sheet to destination workbook
						
						$sh4_wborkbook0 = $WorkBook0.sheets.item(4) # first sheet in destination workbook
						$sheetToCopy4 = $Workbook4.sheets.item(1) # source sheet to copy
						$sheetToCopy4.copy($sh4_wborkbook0) # copy source sheet to destination workbook		

						$sh5_wborkbook0 = $WorkBook0.sheets.item(5) # first sheet in destination workbook
						$sheetToCopy5 = $Workbook5.sheets.item(1) # source sheet to copy
						$sheetToCopy5.copy($sh5_wborkbook0) # copy source sheet to destination workbook									

						$equalboth = $WorkBook0.Worksheets.item(1)
						$missingin1 = $WorkBook0.Worksheets.item(2)
						$missingin2 = $WorkBook0.Worksheets.item(3)
						$diffstate = $WorkBook0.Worksheets.item(4)
						$diffstartmode = $WorkBook0.Worksheets.item(5)
						
						$equalboth.name = 'Same services and values'
						$missingin1.name = 'New services in file 2'
						$missingin2.name = 'New services in file 1'
						$diffstate.name = 'Services with different state'
						$diffstartmode.name = 'Services with diff startmode'
						
						$equalboth.columns.autofit() | out-null
						$missingin1.columns.autofit() | out-null
						$missingin2.columns.autofit() | out-null
						$diffstate.columns.autofit() | out-null
						$diffstartmode.columns.autofit() | out-null
						
						$Table_Equal = $equalboth.ListObjects.add( 1,$equalboth.UsedRange,0,1)	
						$equalboth.ListObjects.Item($Table_Equal.Name).TableStyle="TableStyleMedium6"	

						$Table_Miss1 = $missingin1.ListObjects.add( 1,$missingin1.UsedRange,0,1)	
						$missingin1.ListObjects.Item($Table_Miss1.Name).TableStyle="TableStyleMedium3"	

						$Table_Miss2 = $missingin2.ListObjects.add( 1,$missingin2.UsedRange,0,1)	
						$missingin2.ListObjects.Item($Table_Miss2.Name).TableStyle="TableStyleMedium5"	
						
						$Table_Diffstate = $diffstate.ListObjects.add( 1,$diffstate.UsedRange,0,1)	
						$diffstate.ListObjects.Item($Table_Diffstate.Name).TableStyle="TableStyleMedium8"		

						$Table_Diffstartmode = $diffstartmode.ListObjects.add( 1,$diffstartmode.UsedRange,0,1)	
						$diffstartmode.ListObjects.Item($Table_Diffstartmode.Name).TableStyle="TableStyleMedium9"								

						$WorkBook0.SaveAs($Compare_Services_ToXLS,51)
						$WorkBook0.Saved = $True
						
						$xl.Quit()

						# Check if an Excel process is running with ID different of the $excel_ID_Before ID process
						# If yes it will store IDs in the variable $Excel_Process_After
						# Then all Process in $Excel_Process_After
						$Excel_Process_After = Get-Process Excel | where {$_.id -ne $excel_ID_Before}												
						Foreach ($Process_XL in $Excel_Process_After)						
							{
								stop-process $Process_XL.id	
							}	
					}
			}		

		ElseIf($HTML)
			{											
				$file1_name = split-path $File1 -leaf -resolve	
				$file2_name = split-path $File2 -leaf -resolve	
										
				$Compare_Services_ToHTML = $Path + "\" + "Compare_Services.html"
				
				$nbnewfile1 = 0
				$nbnewfile2 = 0		
				$nbdiffstate = 0
				$nbdiffstartmode = 0		
				$nbsame = 0	
				
				$Title = "<p><span class=Main_Title>Services comparison between $file1_name and $file2_name</span><br><span class=subtitle>This document has been updated on $date</span><br><span class=module_version>Module version: CompareComputer $Module_version</span></p><br><br>"			
				
				### SAME SERVICES BUT DIFFERENT STATE
				$Result_state = "<p class=notequal_list>Services with different state between $file1_name and $file2_name</p>"						
				$Diff_state = ForEach ($line1 in $file1_content)
				{
					ForEach ($line2 in $file2_content)
						{
							IF ($line1.name -eq $line2.name)
								{
									IF ($line1.state -ne $line2.state)
										{		
											New-Object -TypeName PSObject -Property @{
												Name = $line1.name
												"State from file1" = $line1.state
												"State from file2" = $line2.state 												
												}  	
											$nbdiffstate = $nbdiffstate + 1
										}
										Break
								}
						}												
				}

				### SAME SERVICES BUT DIFFERENT STARTMODE
				$Result_startmode = "<p class=notequal_list>Services with different startmode between $file1_name and $file2_name</p>"						
				$Diff_startmode = ForEach ($line1 in $file1_content)
				{
					ForEach ($line2 in $file2_content)
						{
							IF ($line1.name -eq $line2.name)
								{
									IF ($line1.startmode -ne $line2.startmode)
										{		
											New-Object -TypeName PSObject -Property @{
												Name = $line1.name
												"Startmode from file1" = $line1.startmode
												"Startmode from file2" = $line2.startmode 												
												}  	
											$nbdiffstartmode = $nbdiffstartmode + 1	
										}
										Break
								}
						}												
				}

				### NEW SERVICES IN FILE 1
				$Result_newIn1 = "<p class=New_object>New services in $file1_name</p>"						
				$found = $false
				$newin1 = Foreach ($line1 in $file1_content)
					{
						$found = $false
						ForEach ($line2 in $file2_content)
							{
								IF ($line1.name -eq $line2.name)
									{
										$found = $true
										break
									}
							}

						IF (-not $found) 
							{
								New-Object -TypeName PSObject -Property @{
									Name = $line1.name
									State = $line1.state
									Startmode = $line1.startmode 	
									}
								$nbnewfile1 = $nbnewfile1 + 1
							}
					}	

				### NEW SERVICES IN FILE 2
				$Result_newIn2 = "<p class=New_object>New services in $file2_name</p>"						
				$found = $false
				$newin2 = Foreach ($line2 in $file2_content)
					{
						$found = $false
						ForEach ($line1 in $file1_content)
							{
								IF ($line1.name -eq $line2.name)
									{
										$found = $true
										break
									}
							}

						IF (-not $found) 
							{
								New-Object -TypeName PSObject -Property @{
									Name = $line2.name
									State = $line2.startmode
									Startmode = $line2.startmode 	
									}
								$nbnewfile2 = $nbnewfile2 + 1
							}
					}		

				### SAME SERVICES AND SAME VALUES									
					$Same_Values_Title = "<p class=equal_list>Same services and same values</p>"
					$Same_Values = compare-object $file1_content $file2_content -includeequal -property name, state, startmode | Where {$_.SideIndicator -eq "=="} | # | ConvertTo-HTML -Fragment 
					  Group-Object -Property name | % { New-Object psobject -Property @{
						Name=$_.name
						Start_Mode=$_.group[0].startmode
						State=$_.group[0].state
						}}  | Select Name, Start_Mode, State | ConvertTo-HTML -Fragment

				$nbsame = (compare-object $file1_content $file2_content -property name, state, startmode -includeequal  | Where {$_.SideIndicator -eq "=="}|measure-object).count 		
									 
				$Resume_Table =	New-Object -TypeName PSObject -Property @{	
								"Same values" = $nbsame
								"Different state" = $nbdiffstate
								"Different Start Mode" = $nbdiffstartmode 	
								"New services in file2" = $nbnewfile1 								
								"New services in file1" = $nbnewfile2 							
							}
														
				$Resume = $Resume_Table | Select "Same values", "Different state", "Different Start Mode", "New services in file1", "New services in file2" | convertto-html -CSSUri $CSS_File
																								
				# Part to check what to display in the report
				# If there is no same service between both files, the same part will be hidden
				If (($nbsame -eq 0)) 
					{
						$Same_Values_Title = ""
						$Same_Values = ""
					}				
				
				# If there are services with different state and different startmode, both part different state and different startmode will be displayed
				If (($nbdiffstate -ne 0) -and ($nbdiffstartmode -ne 0)) 
					{
						$html1 = $Diff_state | Select Name, "State from file1", "State from file2" | convertto-html -CSSUri $CSS_File				
						$html2 = $Diff_startmode | Select Name, "Startmode from file1", "Startmode from file2"| convertto-html -CSSUri $CSS_File
					}
			
				# If there is no service with different state and no service with different startmode, both part different state and different startmode will be hidden			
				If (($nbdiffstate -eq 0) -and ($nbdiffstartmode -eq 0)) 
					{
						$Result_state = ""					
						$Result_startmode = ""					
					}
					
				# If there is no service with different state but there are service with different startmode, only the part different startmode will be displayed	
				If (($nbdiffstate -eq 0) -and ($nbdiffstartmode -ne 0)) 
					{
						$Result_state = ""		
						$html2 = $Diff_startmode | Select Name, "Startmode from file1", "Startmode from file2"| convertto-html -CSSUri $CSS_File						
					}	
					
				# If there are services with different state but there is no service with different startmode, only the part different state will be displayed						
				If (($nbdiffstate -ne 0) -and ($nbdiffstartmode -eq 0)) 
					{					
						$Result_startmode = ""	
						$html1 = $Diff_state | Select Name, "State from file1", "State from file2" | convertto-html -CSSUri $CSS_File										
					}		
				
				# If there are new services in file 1 and file 2, both new services in file1 and new services in file2 parts will be displayed
				If (($nbnewfile1 -ne 0) -and ($nbnewfile2 -ne 0)) 
					{
						$html3 = $newin1 | Select Name, Startmode, State | convertto-html -CSSUri $CSS_File
						$html4 = $newin2 | Select Name, Startmode, State | convertto-html -CSSUri $CSS_File
					}
				
				# If there is no new services in file 1 and no new services in file 2, both new services in file1 and new services in file2 parts will be hidden
				If (($nbnewfile1 -eq 0) -and ($nbnewfile2 -eq 0)) 
					{
						$Result_newIn1 = ""
						$Result_newIn2 = ""
					}					
										
				# If there is no new services in file 2 and if there are new services in file 1, the new services in part 1 will be displayed and the new services in file 2 will be hidden										
				If (($nbnewfile1 -ne 0) -and ($nbnewfile2 -eq 0)) 
					{
						$Result_newIn2 = ""
						$html3 = $newin1 | Select Name, Startmode, State | convertto-html -CSSUri $CSS_File						
					}	

				# If there is no new services in file 1 and if there are new services in file 2, the new services in part 2 will be displayed and the new services in file 1 will be hidden															
				If (($nbnewfile1 -eq 0) -and ($nbnewfile2 -ne 0)) 
					{
						$Result_newIn1 = ""	
						$html3 = $newin2 | Select Name, Startmode, State | convertto-html -CSSUri $CSS_File						
					}							

				$html_final = convertto-html -body "$Title<span class=Resume_Title>Resume values</span><br><br>$Resume <br><br><br>				
				
				<div id=left>$Same_Values_Title $Same_Values</div>
				<div id=right_services>$Result_state $html1 
					<br> 
					$Result_startmode $html2
					<br>
					$Result_newIn1 $html3
					<br><br> 
					$Result_newIn2 $html4
				</div>						
				" -CSSUri $CSS_File

				$html_final | out-file -encoding ASCII $Compare_Services_ToHTML
				invoke-expression $Compare_Services_ToHTML					
			}				
	}

    end
    {
		If($XLS)
			{			
				If ($Excel_value -eq $True)
					{			
						write-host "Services have been compared in XLS format" -foregroundcolor "Cyan"
						write-host "********************************************************************************************************"	
						If (($nbnewfile1 -ne 0) -or ($nbnewfile2 -ne 0) -or ($nbdiffstate -ne 0) -or ($nbdiffstartmode -ne 0))
							{
								write-host "Comparison Status: Some services services seems to be different between the two files !!!" -foregroundcolor "yellow"	 							
							}
						ElseIf (($nbnewfile1 -eq 0) -or ($nbnewfile2 -eq 0) -or ($nbdiffstate -eq 0) -or ($nbdiffstartmode -eq 0))
							{
								write-host "Comparison Status: All services are similar !!!" -foregroundcolor "green"	 							
							}	
							
						write-host "********************************************************************************************************"	
						write-host "See below the results of the comparison" -foregroundcolor "Cyan"
						write-host ""							
						write-host "Services with equal values : $nbsame"
						write-host "New services in file 1 : $nbnewfile1"
						write-host "New services in file 2 : $nbnewfile2"
						write-host "Services with different state : $nbdiffstate"		
						write-host "Services with different startmode : $nbdiffstartmode"								
						write-host "********************************************************************************************************"	
					}
				Else
					{
						write-host "Excel seems to be not installed"	
					}			
			}					
			
		ElseIf($HTML)
			{										
				write-host "Services have been compared in HTML format" -foregroundcolor "Cyan"
				write-host "********************************************************************************************************"	
				If (($nbnewfile1 -ne 0) -or ($nbnewfile2 -ne 0) -or ($nbdiffstate -ne 0) -or ($nbdiffstartmode -ne 0))
					{
						write-host "Comparison Status: Some services services seems to be different between the two files !!!" -foregroundcolor "yellow"	 							
					}
				ElseIf (($nbnewfile1 -eq 0) -and ($nbnewfile2 -eq 0) -and ($nbdiffstate -eq 0) -and ($nbdiffstartmode -eq 0))
					{
						write-host "Comparison Status: All services are similar !!!" -foregroundcolor "green"	 							
					}	
					
				write-host "********************************************************************************************************"	
				write-host "See below the results of the comparison" -foregroundcolor "Cyan"
				write-host ""							
				write-host "Services with equal values : $nbsame"
				write-host "New services in file 1 : $nbnewfile1"
				write-host "New services in file 2 : $nbnewfile2"
				write-host "Services with different state : $nbdiffstate"		
				write-host "Services with different startmode : $nbdiffstartmode"								
				write-host "********************************************************************************************************"									
			}
	}								
}










		
		
			
	

						



	
	
			

					


						







































<#.Synopsis
	The Compare-Process function allows you to compare two computer configurations using two CSV or XML files.
	It works like this: Compare-Process -Path C:\ -file1 <File1_Path> -file2 <File2_Path> -CSV or -XLS
.DESCRIPTION
	Allow you to compare two computer configurations by comparing twi CSV or XML files.
	These CSV or XML files have to contain bith Process export with Name, Caption, State and Startmode columns
	They could be generated easily using the ConfigExport module
	Comparison can be done in two formats: HTML or XLSX

.EXAMPLE
PS Root\> Compare-Process -Path C:\ -file1 c:\Process1.csv -file2 c:\Process2.csv -HTML
The command above will compare both Process1.csv and Process2.csv files. 
It will create a file Compare_Process.html file in C:

.EXAMPLE
PS Root\> Compare-Process -Path C:\ -file1 c:\Process1.csv -file2 c:\Process2.csv -XLS
The command above will compare both Process1.csv and Process2.csv files. 
It will create a file Compare_Process.xlsx file in C:

.NOTES
    Author: Damien VAN ROBAEYS - @syst_and_deploy - http://www.systanddeploy.com
#>

Function Compare-Process
{
[CmdletBinding()]
Param(
	[Parameter(Mandatory=$true,ValueFromPipeline=$true, position=1)]
	[string] $Path,
	[Parameter(Mandatory=$true)]		
	[string] $File1,
	[Parameter(Mandatory=$true)]		
	[string] $File2,		
	[Switch] $XLS,
	[Switch] $HTML				
    )
    
    Begin
    {		
		# If both files File1 and File2 are CSV
		If (($File1.contains("csv")) -and ($File2.contains("csv")))
			{
				$file1_content = import-csv $File1	 
				$file2_content = import-csv $File2							
			}
			
		# If both files File1 and File2 are XML			
		ElseIf (($File1.contains("xml")) -and ($File2.contains("xml")))
			{
				$file1_content = import-clixml $File1	 
				$file2_content = import-clixml $File2							
			}	
	
		# If File1 is XML and File2 is CSV			
		ElseIf (($File1.contains("xml")) -and ($File2.contains("csv")))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You have specified two different files to compare" -foregroundcolor "yellow"		
				write-host " !!! The first file is a XML and the second file is an CSV" -foregroundcolor "yellow"		 								
				write-host " !!! You need to specify two CSV or two XML files for the comparison" -foregroundcolor "yellow"										
				write-host "*******************************************************************"			
				break
			}
			
		# If File1 is CSV and File2 is XML						
		ElseIf (($File1.contains("csv")) -and ($File2.contains("xml")))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You have specified two different files to compare" -foregroundcolor "yellow"		
				write-host " !!! The first file is a CSV and the second file is an XML" -foregroundcolor "yellow"	
				write-host " !!! You need to specify two CSV or two XML files for the comparison" -foregroundcolor "yellow"										
				write-host "*******************************************************************"		
				break				
			}		
	
		# If there is no output format specified
		If ((-not $XLS) -and (-not $HTML))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You need to specific an output format. " -foregroundcolor "yellow"		
				write-host " !!! Use the switch -XLS to export in CSV or XLS format " -foregroundcolor "yellow"		 	
				write-host " !!! Use the switch -html to export in HTML format " -foregroundcolor "yellow"										
				write-host "*******************************************************************"					
			}	

		# If there output format is XLS				
		If($XLS)
			{			
				Try
					{
						# Check if an Excel process is already running, if yes it will keep the ID in the variable $excel_ID_Before
						$Excel_Process_Before = Get-Process Excel -ErrorAction SilentlyContinue
						$excel_ID_Before = $Excel_Process_Before.id						
										
						$Excel_Test = new-object -comobject excel.application
						$Excel_value = $True 
						write-host ""		
						write-host "********************************************************************************************************"	
						write-host "Process will be compared in XLS format" -foregroundcolor "Cyan"						
					}

				Catch 
					{
						$Excel_value = $false 					
					}		
			}	
			
		# If there output format is HTML			
		ElseIf($HTML)
			{
				write-host ""		
				write-host "********************************************************************************************************"	
				write-host "Process will be exported in HTML format" -foregroundcolor "Cyan"				
			}				
    }

	
    Process
    {
		If($XLS)
			{				
				If ($Excel_value -eq $True)				
					{						
						$file1_name = split-path $File1 -leaf -resolve	
						$file2_name = split-path $File2 -leaf -resolve	
						
						$Compare_Process_ToXLS = $Path + "\" + "Compare_Process.xlsx"
																			
						$nbnewfile1 = 0
						$nbnewfile2 = 0		
						$nbsame = 0	
						
						$equal = compare-object $file1_content $file2_content -property name -includeequal  | Where {$_.SideIndicator -eq "=="} | 
						   Group-Object -Property name | % { New-Object psobject -Property @{		
								Name=$_.group[0].name
							}}  | export-csv -encoding UTF8 -notype  $Temp_CSV_Equal 
										
						$missingfile1 = compare-object $file1_content $file2_content -property name, executablepath  | Where {$_.SideIndicator -eq "=>"} | 
						   Group-Object -Property name | % { New-Object psobject -Property @{		
								Name=$_.group[0].name
							}}  | export-csv -encoding UTF8 -notype  $Temp_CSV_MissingInFile1

						$missingfile2 = compare-object $file1_content $file2_content -property name, executablepath  | Where {$_.SideIndicator -eq "<="} | 
						   Group-Object -Property name | % { New-Object psobject -Property @{		
								Name=$_.group[0].name
							}}  | export-csv -encoding UTF8 -notype  $Temp_CSV_MissingInFile2


						$nbsame = (compare-object $file1_content $file2_content -property name -includeequal  | Where {$_.SideIndicator -eq "=="} | 
						   Group-Object -Property name | % { New-Object psobject -Property @{		
								Name=$_.group[0].name
								Description=$_.group[0].executablepath	
							}}  |measure-object).count	

						$nbnewfile2 = (compare-object $file1_content $file2_content -property name, executablepath  | Where {$_.SideIndicator -eq "=>"} | 
						   Group-Object -Property name | % { New-Object psobject -Property @{		
								Name=$_.group[0].name
							}}  |measure-object).count	

						$nbnewfile1 = (compare-object $file1_content $file2_content -property name, executablepath  | Where {$_.SideIndicator -eq "<="} | 
						   Group-Object -Property name | % { New-Object psobject -Property @{		
								Name=$_.group[0].name
							}}  |measure-object).count				
														
						$xl = new-object -comobject excel.application
						$xl.visible = $false
						$xl.DisplayAlerts=$False

						$Workbook1 = $xl.workbooks.open($Temp_CSV_Equal)
						$Workbook2 = $xl.workbooks.open($Temp_CSV_MissingInFile1)
						$Workbook3 = $xl.workbooks.open($Temp_CSV_MissingInFile2) 

						$WorkBook0 = $xl.WorkBooks.add()

						$sh1_wborkbook0 = $WorkBook0.sheets.item(1) # first sheet in destination workbook
						$sheetToCopy1 = $Workbook1.sheets.item(1) # source sheet to copy
						$sheetToCopy1.copy($sh1_wborkbook0) # copy source sheet to destination workbook

						$sh2_wborkbook0 = $WorkBook0.sheets.item(2) # first sheet in destination workbook
						$sheetToCopy2 = $Workbook2.sheets.item(1) # source sheet to copy
						$sheetToCopy2.copy($sh2_wborkbook0) # copy source sheet to destination workbook

						$sh3_wborkbook0 = $WorkBook0.sheets.item(3) # first sheet in destination workbook
						$sheetToCopy3 = $Workbook3.sheets.item(1) # source sheet to copy
						$sheetToCopy3.copy($sh3_wborkbook0) # copy source sheet to destination workbook

						$equalboth = $WorkBook0.Worksheets.item(1)
						$missingin1 = $WorkBook0.Worksheets.item(2)
						$missingin2 = $WorkBook0.Worksheets.item(3)

						$equalboth.name = 'Same process'
						$missingin1.name = 'New process in file 2'
						$missingin2.name = 'New process in file 1'

						$equalboth.columns.autofit() | out-null
						$missingin1.columns.autofit() | out-null
						$missingin2.columns.autofit() | out-null

						$Table_Equal = $equalboth.ListObjects.add( 1,$equalboth.UsedRange,0,1)	
						$equalboth.ListObjects.Item($Table_Equal.Name).TableStyle="TableStyleMedium6"	

						$Table_Miss1 = $missingin1.ListObjects.add( 1,$missingin1.UsedRange,0,1)	
						$missingin1.ListObjects.Item($Table_Miss1.Name).TableStyle="TableStyleMedium3"	

						$Table_Miss1 = $missingin2.ListObjects.add( 1,$missingin2.UsedRange,0,1)	
						$missingin2.ListObjects.Item($Table_Miss1.Name).TableStyle="TableStyleMedium5"	

						$WorkBook0.SaveAs($Compare_Process_ToXLS,51)
						$WorkBook0.Saved = $True
						$xl.Quit()	

						# Check if an Excel process is running with ID different of the $excel_ID_Before ID process
						# If yes it will store IDs in the variable $Excel_Process_After
						# Then all Process in $Excel_Process_After
						$Excel_Process_After = Get-Process Excel | where {$_.id -ne $excel_ID_Before}												
						Foreach ($Process_XL in $Excel_Process_After)						
							{
								stop-process $Process_XL.id	
							}																								
					}						
			}		

		ElseIf($HTML)
			{
				$file1_name = split-path $File1 -leaf -resolve	
				$file2_name = split-path $File2 -leaf -resolve	
						
				$Compare_Process_ToHTML = $Path + "\" + "Compare_Process.html"
								
				$nbnewfile1 = 0
				$nbnewfile2 = 0		
				$nbsame = 0
				
				$Title = "<p><span class=Main_Title>Process comparison between $file1_name and $file2_name</span><br><span class=subtitle>This document has been updated on $date</span><br><span class=module_version>Module version: CompareComputer $Module_version</span></p><br><br>"							
				
				# Part where process are equals
				$Same_Values_Title = "<p class=equal_list>Same Process</p>"
				$Same_Values = compare-object $file1_content $file2_content -includeequal -property name | Where {$_.SideIndicator -eq "=="}  | 
					Group-Object -Property name | % { New-Object psobject -Property @{
						Process_Name=$_.group[0].name	
					}}  | ConvertTo-HTML -Fragment  
													 
				# Part where process are missing in $file1
				$Title_NewInfile1 = "<p class=New_object>New process in $file1_name</p>"
				$NewInfile1 = compare-object $file1_content $file2_content -property name | Where {$_.SideIndicator -eq "=>"} | 
				   Group-Object -Property name | % { New-Object psobject -Property @{		
						Process_Name=$_.group[0].name	
					}}  | ConvertTo-HTML -Fragment 
					
				# Part where process are missing in $file2
				$Title_NewInfile2 = "<p class=New_object>New process in $file2_name</p>"
				$NewInfile2 = compare-object $file1_content $file2_content -property name | Where {$_.SideIndicator -eq "<="} |
					Group-Object -Property name | % { New-Object psobject -Property @{
						Process_Name=$_.group[0].name	
					}}  | ConvertTo-HTML -Fragment 


				$nbsame = (compare-object $file1_content $file2_content -includeequal -property name | Where {$_.SideIndicator -eq "=="}  | 
					Group-Object -Property name | % { New-Object psobject -Property @{
						Process_Name=$_.group[0].name	
					}} |measure-object).count			

				$nbnewfile2 = (compare-object $file1_content $file2_content -property name | Where {$_.SideIndicator -eq "=>"} | 
				   Group-Object -Property name | % { New-Object psobject -Property @{		
						Process_Name=$_.group[0].name	
					}}  |measure-object).count 


				$nbnewfile1 = (compare-object $file1_content $file2_content -property name | Where {$_.SideIndicator -eq "<="} |
					Group-Object -Property name | % { New-Object psobject -Property @{
						Process_Name=$_.group[0].name	
					}}  |measure-object).count  
																						
				$Resume_Table =	New-Object -TypeName PSObject -Property @{	
								"Same process" = $nbsame
								"New process in file 1" = $nbnewfile1
								"New process in file 2" = $nbnewfile2 	
							}
							
				$Resume = $Resume_Table | Select "Same process", "New process in file 1", "New process in file 2" | convertto-html -CSSUri $CSS_File
																									
				# Part to check what to display in the report
				
				# If there is no same process between both files, the same part will be hidden
				If (($nbsame -eq 0)) 
					{
						$Same_Values_Title = ""
						$Same_Values = ""
					}	

				# If there are new process in file 1 and 2, both parts new processin file 1 and new process in file 2 will be displayed
				If (($nbnewfile1 -ne 0) -and ($nbnewfile2 -ne 0)) 
					{
					}
			
				# If there is no new process in file 1 and 2, both parts new process in file 1 and new process in file 2 will be hidden 
				If (($nbnewfile1 -eq 0) -and ($nbnewfile2 -eq 0)) 
					{
						$Title_NewInfile1 = ""
						$Title_NewInfile2 = ""
						$NewInfile1 = ""						
						$NewInfile2 = ""
					}
					
				# If there is no new process in file 1 and there are new process in file 2, the part new process in file 1 part will be hidden and the part new process in file 2 will be displayed	
				If (($nbnewfile1 -eq 0) -and ($nbnewfile2 -ne 0)) 
					{
						$Title_NewInfile1 = ""
						$NewInfile1 = ""						
					}	
					
				# If there is no new process in file 2 and there are new process in file 1, the part new process in file 2 part will be hidden and the part new process in file 1 will be displayed	
				If (($nbnewfile1 -ne 0) -and ($nbnewfile2 -eq 0)) 
					{		
						$Title_NewInfile2 = ""
						$NewInfile2 = ""						
					}		
								
				$html_final = convertto-html -body "$Title<span class=Resume_Title>Resume values</span><br><br>$Resume <br><br>				
				<div id=left>$Same_Values_Title $Same_Values</div>
				<div id=right_process>
					$Title_NewInfile1 $NewInfile1
					<br>
					$Title_NewInfile2 $NewInfile2								
				</div></center>					
				" -CSSUri $CSS_File

				$html_final | out-file -encoding ASCII $Compare_Process_ToHTML
				invoke-expression $Compare_Process_ToHTML														
			}				
	}	
	

    end
    {			
		If($XLS)
			{
				write-host "Process have been compared in HTML format" -foregroundcolor "Cyan"
				write-host "********************************************************************************************************"	
				If (($nbnewfile1 -ne 0) -or ($nbnewfile2 -ne 0))
					{
						write-host "Comparison Status: Some Process seems to be different between the two files !!!" -foregroundcolor "yellow"	 							
					}
				ElseIf (($nbnewfile1 -eq 0) -and ($nbnewfile2 -eq 0))
					{
						write-host "Comparison Status: All Process are similar !!!" -foregroundcolor "green"	 							
					}	
					
				write-host "********************************************************************************************************"	
				write-host "See below the results of the comparison" -foregroundcolor "Cyan"
				write-host ""							
				write-host "Process with equal values : $nbsame"
				write-host "New process in file 1 : $nbnewfile1"
				write-host "New process in file 2 : $nbnewfile2"
				write-host "********************************************************************************************************"				
			}	
			
		ElseIf($HTML)			
			{										
				write-host "Process have been compared in HTML format" -foregroundcolor "Cyan"
				write-host "********************************************************************************************************"	
				If (($nbnewfile1 -ne 0) -or ($nbnewfile2 -ne 0))
					{
						write-host "Comparison Status: Some Process seems to be different between the two files !!!" -foregroundcolor "yellow"	 							
					}
				ElseIf (($nbnewfile1 -eq 0) -and ($nbnewfile2 -eq 0))
					{
						write-host "Comparison Status: All Process are similar !!!" -foregroundcolor "green"	 							
					}	
					
				write-host "********************************************************************************************************"	
				write-host "See below the results of the comparison" -foregroundcolor "Cyan"
				write-host ""							
				write-host "Process with equal values : $nbsame"
				write-host "New process in file 1 : $nbnewfile1"
				write-host "New process in file 2 : $nbnewfile2"
				write-host "********************************************************************************************************"											
			}								
	}
	
}











<#.Synopsis
	The Compare-Hotfix function allows you to compare two computer configurations using two CSV or XML files.
	It works like this: Compare-Hotfix -Path C:\ -file1 <File1_Path> -file2 <File2_Path> -CSV or -XLS
.DESCRIPTION
	Allow you to compare two computer configurations by comparing twi CSV or XML files.
	These CSV or XML files have to contain bith Hotfix export with Hotfixid and Description columns
	They could be generated easily using the ConfigExport module
	Comparison can be done in two formats: HTML or XLSX

.EXAMPLE
PS Root\> Compare-Hotfix -Path C:\ -file1 c:\Hotfix1.csv -file2 c:\Hotfix2.csv -HTML
The command above will compare both Hotfix1.csv and Hotfix2.csv files. 
It will create a file Compare_Hotfix.html file in C:

.EXAMPLE
PS Root\> Compare-Hotfix -Path C:\ -file1 c:\Process1.csv -file2 c:\Process2.csv -XLS
The command above will compare both Hotfix1.csv and Hotfix2.csv files. 
It will create a file Compare_Hotfix.xlsx file in C:

.NOTES
    Author: Damien VAN ROBAEYS - @syst_and_deploy - http://www.systanddeploy.com
#>

Function Compare-Hotfix
{
[CmdletBinding()]
Param(
	[Parameter(Mandatory=$true,ValueFromPipeline=$true, position=1)]
	[string] $Path,
	[Parameter(Mandatory=$true)]		
	[string] $File1,
	[Parameter(Mandatory=$true)]		
	[string] $File2,		
	[Switch] $XLS,
	[Switch] $HTML				
    )


    Begin
    {		
		# If both files File1 and File2 are CSV
		If (($File1.contains("csv")) -and ($File2.contains("csv")))
			{
				$file1_content = import-csv $File1	 
				$file2_content = import-csv $File2							
			}
			
		# If both files File1 and File2 are XML			
		ElseIf (($File1.contains("xml")) -and ($File2.contains("xml")))
			{
				$file1_content = import-clixml $File1	 
				$file2_content = import-clixml $File2							
			}	
	
		# If File1 is XML and File2 is CSV			
		ElseIf (($File1.contains("xml")) -and ($File2.contains("csv")))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You have specified two different files to compare" -foregroundcolor "yellow"		
				write-host " !!! The first file is a XML and the second file is an CSV" -foregroundcolor "yellow"		 								
				write-host " !!! You need to specify two CSV or two XML files for the comparison" -foregroundcolor "yellow"										
				write-host "*******************************************************************"			
				break
			}
			
		# If File1 is CSV and File2 is XML						
		ElseIf (($File1.contains("csv")) -and ($File2.contains("xml")))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You have specified two different files to compare" -foregroundcolor "yellow"		
				write-host " !!! The first file is a CSV and the second file is an XML" -foregroundcolor "yellow"	
				write-host " !!! You need to specify two CSV or two XML files for the comparison" -foregroundcolor "yellow"										
				write-host "*******************************************************************"		
				break				
			}			
	
		# If there is no output format specified
		If ((-not $XLS) -and (-not $HTML))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You need to specific an output format. " -foregroundcolor "yellow"		
				write-host " !!! Use the switch -XLS to export in CSV or XLS format " -foregroundcolor "yellow"		 	
				write-host " !!! Use the switch -html to export in HTML format " -foregroundcolor "yellow"										
				write-host "*******************************************************************"					
			}		

		# If there output format is XLS					
		If($XLS)
			{			
				Try
					{
						# Check if an Excel process is already running, if yes it will keep the ID in the variable $excel_ID_Before
						$Excel_Process_Before = Get-Process Excel -ErrorAction SilentlyContinue
						$excel_ID_Before = $Excel_Process_Before.id					
										
						$Excel_Test = new-object -comobject excel.application
						$Excel_value = $True 
						write-host ""		
						write-host "********************************************************************************************************"	
						write-host "Hotfix will be compared in XLS format" -foregroundcolor "Cyan"					
					}

				Catch 
					{
						$Excel_value = $false 					
					}		
			}	
			
		# If there output format is HTML			
		ElseIf($HTML)
			{
				$CSS_File = "$Current_Folder\Master_Export_Compare.css" # CSS for HTML Export
				write-host ""		
				write-host "********************************************************************************************************"	
				write-host "Hotfix will be exported in HTML format" -foregroundcolor "Cyan"													
			}				
    }

	
    Process
    {
		If($XLS)
			{
						
				If ($Excel_value -eq $True)				
					{				
						$file1_name = split-path $File1 -leaf -resolve	
						$file2_name = split-path $File2 -leaf -resolve	
								
						$Compare_Hotfix_ToXLS = $Path + "\" + "Compare_Hotfix.xlsx"
												
						$nbnewfile1 = 0
						$nbnewfile2 = 0		
						$nbsame = 0	
						
						$equal = compare-object $file1_content $file2_content -property hotfixid, description -includeequal  | Where {$_.SideIndicator -eq "=="} | 
						   Group-Object -Property hotfixid | % { New-Object psobject -Property @{		
								Hotfix_ID=$_.group[0].hotfixid
								Description=$_.group[0].description	
							}}  | export-csv -encoding UTF8 -notype  $Temp_CSV_Equal 
										
						$missingfile1 = compare-object $file1_content $file2_content -property hotfixid, description  | Where {$_.SideIndicator -eq "=>"} | 
						   Group-Object -Property hotfixid | % { New-Object psobject -Property @{		
								Hotfix_ID=$_.group[0].hotfixid
								Description=$_.group[0].description	
							}}  | export-csv -encoding UTF8 -notype  $Temp_CSV_MissingInFile1

						$missingfile2 = compare-object $file1_content $file2_content -property hotfixid, description  | Where {$_.SideIndicator -eq "<="} | 
						   Group-Object -Property hotfixid | % { New-Object psobject -Property @{		
								Hotfix_ID=$_.group[0].hotfixid
								Description=$_.group[0].description	
							}}  | export-csv -encoding UTF8 -notype  $Temp_CSV_MissingInFile2

							
						$nbsame = (compare-object $file1_content $file2_content -property hotfixid, description -includeequal  | Where {$_.SideIndicator -eq "=="}|measure-object).count 		
						$nbnewfile1 = (compare-object $file1_content $file2_content -property hotfixid, description  | Where {$_.SideIndicator -eq "<="}|measure-object).count 		
						$nbnewfile2 = (compare-object $file1_content $file2_content -property hotfixid, description  | Where {$_.SideIndicator -eq "=>"}|measure-object).count 		
												
						$xl = new-object -comobject excel.application
						$xl.visible = $false
						$xl.DisplayAlerts=$False

						$Workbook1 = $xl.workbooks.open($Temp_CSV_Equal)
						$Workbook2 = $xl.workbooks.open($Temp_CSV_MissingInFile1)
						$Workbook3 = $xl.workbooks.open($Temp_CSV_MissingInFile2) 

						$WorkBook0 = $xl.WorkBooks.add()

						$sh1_wborkbook0 = $WorkBook0.sheets.item(1) # first sheet in destination workbook
						$sheetToCopy1 = $Workbook1.sheets.item(1) # source sheet to copy
						$sheetToCopy1.copy($sh1_wborkbook0) # copy source sheet to destination workbook

						$sh2_wborkbook0 = $WorkBook0.sheets.item(2) # first sheet in destination workbook
						$sheetToCopy2 = $Workbook2.sheets.item(1) # source sheet to copy
						$sheetToCopy2.copy($sh2_wborkbook0) # copy source sheet to destination workbook

						$sh3_wborkbook0 = $WorkBook0.sheets.item(3) # first sheet in destination workbook
						$sheetToCopy3 = $Workbook3.sheets.item(1) # source sheet to copy
						$sheetToCopy3.copy($sh3_wborkbook0) # copy source sheet to destination workbook

						$equalboth = $WorkBook0.Worksheets.item(1)
						$missingin1 = $WorkBook0.Worksheets.item(2)
						$missingin2 = $WorkBook0.Worksheets.item(3)

						$equalboth.name = 'Same Hotfix'
						$missingin1.name = 'New Hotfix in file 2'
						$missingin2.name = 'New Hotfix in file 1'
						
						$equalboth.columns.autofit() | out-null
						$missingin1.columns.autofit() | out-null
						$missingin2.columns.autofit() | out-null

						$Table_Equal = $equalboth.ListObjects.add( 1,$equalboth.UsedRange,0,1)	
						$equalboth.ListObjects.Item($Table_Equal.Name).TableStyle="TableStyleMedium6"	

						$Table_Miss1 = $missingin1.ListObjects.add( 1,$missingin1.UsedRange,0,1)	
						$missingin1.ListObjects.Item($Table_Miss1.Name).TableStyle="TableStyleMedium3"	

						$Table_Miss1 = $missingin2.ListObjects.add( 1,$missingin2.UsedRange,0,1)	
						$missingin2.ListObjects.Item($Table_Miss1.Name).TableStyle="TableStyleMedium5"	

						$WorkBook0.SaveAs($Compare_Hotfix_ToXLS,51)
						$WorkBook0.Saved = $True
						$xl.Quit()		

						# Check if an Excel process is running with ID different of the $excel_ID_Before ID process
						# If yes it will store IDs in the variable $Excel_Process_After
						# Then all Process in $Excel_Process_After
						$Excel_Process_After = Get-Process Excel | where {$_.id -ne $excel_ID_Before}												
						Foreach ($Process_XL in $Excel_Process_After)						
							{
								stop-process $Process_XL.id	
							}						
					}									
			}		
		ElseIf($HTML)
			{		
				$file1_name = split-path $File1 -leaf -resolve	
				$file2_name = split-path $File2 -leaf -resolve	
				
				$Compare_Hotfix_ToHTML = $Path + "\" + "Compare_Hotfix.html"
				
				$nbnewfile1 = 0
				$nbnewfile2 = 0		
				$nbsame = 0		
			
				$Title = "<p><span class=Main_Title>Hotfix comparison between $file1_name and $file2_name</span><br><span class=subtitle>This document has been updated on $date</span><br><span class=module_version>Module version: CompareComputer $Module_version</span></p><br><br>"			
	
				# Part where Hotfix are equals
				$Same_Values_Title = "<p class=equal_list>Same hotfix</p>"
				$Same_Values = compare-object $file1_content $file2_content -includeequal -property hotfixid, description | Where {$_.SideIndicator -eq "=="}  | #ConvertTo-HTML -Fragment 
					Group-Object -Property hotfixid | % { New-Object psobject -Property @{
						Hotfix_ID=$_.group[0].hotfixid	
						Description=$_.group[0].description			
					}}  | ConvertTo-HTML -Fragment  
					
				#Part where Hotfix are not in $file1
				$Title_NewInfile1 = "<p class=New_object>New Hotfix in $file1_name</p>"
				$NewInfile1 = compare-object $file1_content $file2_content -property hotfixid, description | Where {$_.SideIndicator -eq "=>"} | #ConvertTo-HTML -Fragment 
				   Group-Object -Property hotfixid | % { New-Object psobject -Property @{		
						Hotfix_ID=$_.group[0].hotfixid
						Description=$_.group[0].description	
					}}  | ConvertTo-HTML -Fragment 
					
				#Part where Hotfix are not in $file2
				$Title_NewInfile2 = "<p class=New_object>New Hotfix in $file2_name</p>"
				$NewInfile2 = compare-object $file1_content $file2_content -property hotfixid, description | Where {$_.SideIndicator -eq "<="} | #ConvertTo-HTML -Fragment 
					Group-Object -Property hotfixid | % { New-Object psobject -Property @{
						Hotfix_ID=$_.group[0].hotfixid
						Description=$_.group[0].description
					}}  | ConvertTo-HTML -Fragment 

							
				$nbsame = (compare-object $file1_content $file2_content -property hotfixid, description -includeequal  | Where {$_.SideIndicator -eq "=="}|measure-object).count 	
				$nbnewfile1 = (compare-object $file1_content $file2_content -property hotfixid, description  | Where {$_.SideIndicator -eq "=>"}|measure-object).count 						
				$nbnewfile2 = (compare-object $file1_content $file2_content -property hotfixid, description  | Where {$_.SideIndicator -eq "<="}|measure-object).count 		
						
				$Resume_Table =	New-Object -TypeName PSObject -Property @{	
								"Same Hotfix" = $nbsame
								"New Hotfix in file 1" = $nbnewfile1
								"New Hotfix in file 2" = $nbnewfile2 	
							}
				
				$Resume = $Resume_Table | Select "Same Hotfix", "New Hotfix in file 1", "New Hotfix in file 2" | convertto-html -CSSUri $CSS_File
									
				# Part to check what to display in the report
				
				# If there is no same hotfix between files, the part same hotfix will be hidden
				If (($nbsame -eq 0)) 
					{
						$Same_Values_Title = ""
						$Same_Values = ""
					}					
				
				# If there are new hotfix in file 1 and file 2, both parts new hotfix in file 1 and new hotfix in file 2 will be displayed
				If (($nbnewfile1 -ne 0) -and ($nbnewfile2 -ne 0)) 
					{
					}
			
				# If there is no new hotfix in file 1 and file 2, both parts new hotfix in file 1 and new hotfix in file 2 will be hidden
				If (($nbnewfile1 -eq 0) -and ($nbnewfile2 -eq 0)) 
					{
						$Title_NewInfile1 = ""
						$Title_NewInfile2 = ""
						$NewInfile1 = ""						
						$NewInfile2 = ""
					}
					
				# If there is no new hotfix in file 1 and there are new hotfix in file 2, the part new hotfix in file 1 will be hidden and the part new hotfix in file 2 will be displayed	
				If (($nbnewfile1 -eq 0) -and ($nbnewfile2 -ne 0)) 
					{
						$Title_NewInfile1 = ""
						$NewInfile1 = ""						
					}	
					
				# If there is no new hotfix in file 2 and there are new hotfix in file 1, the part new hotfix in file 2 will be hidden and the part new hotfix in file 1 will be displayed	
				If (($nbnewfile1 -ne 0) -and ($nbnewfile2 -eq 0)) 
					{		
						$Title_NewInfile2 = ""
						$NewInfile2 = ""						
					}			

				$html_final = convertto-html -body "$Title<span class=Resume_Title>Resume values</span><br><br>$Resume <br><br>				
				<div id=left>$Same_Values_Title<br>$Same_Values</div>
				<div id=right_process>
					$Title_NewInfile1 $NewInfile1
					<br>
					$Title_NewInfile2 $NewInfile2								
				</div></center>					
				" -CSSUri $CSS_File	
				$html_final | out-file -encoding ASCII $Compare_Hotfix_ToHTML
				invoke-expression $Compare_Hotfix_ToHTML				
			}		
	}	

    end
    {
		If($XLS)
			{			
				If ($Excel_value -eq $True)
					{											
						write-host "Hotfix have been compared in XLS format" -foregroundcolor "Cyan"
						write-host "********************************************************************************************************"	
						If (($nbnewfile1 -ne 0) -or ($nbnewfile2 -ne 0))
							{
								write-host "Comparison Status: Some Hotfix seems to be different between the two files !!!" -foregroundcolor "yellow"	 							
							}
						ElseIf (($nbnewfile1 -eq 0) -and ($nbnewfile2 -eq 0))
							{
								write-host "Comparison Status: All Hotfix are similar !!!" -foregroundcolor "green"	 							
							}	
							
						write-host "********************************************************************************************************"	
						write-host "See below the results of the comparison" -foregroundcolor "Cyan"
						write-host ""							
						write-host "Hotfix with equal values : $nbsame"
						write-host "New Hotfix in file 1 : $nbnewfile1"
						write-host "New Hotfix in file 2 : $nbnewfile2"
						write-host "********************************************************************************************************"	
					}
				Else
					{
						write-host "Excel seems to be not installed"	
					}
			}				
			
		ElseIf($HTML)
			{										
				write-host "Hotfix have been compared in HTML format" -foregroundcolor "Cyan"
				write-host "********************************************************************************************************"	
				If (($nbnewfile1 -ne 0) -or ($nbnewfile2 -ne 0))
					{
						write-host "Comparison Status: Some Hotfix seems to be different between the two files !!!" -foregroundcolor "yellow"	 							
					}
				ElseIf (($nbnewfile1 -eq 0) -and ($nbnewfile2 -eq 0))
					{
						write-host "Comparison Status: All Hotfix are similar !!!" -foregroundcolor "green"	 							
					}	
					
						write-host "********************************************************************************************************"	
						write-host "See below the results of the comparison" -foregroundcolor "Cyan"
						write-host ""							
						write-host "Hotfix with equal values : $nbsame"
						write-host "New Hotfix in file 1 : $nbnewfile1"
						write-host "New Hotfix in file 2 : $nbnewfile2"
						write-host "********************************************************************************************************"																		
			}
	}
	
}

























<#.Synopsis
	The Compare-Drivers function allows you to export a Drivers list from your computer. 
.DESCRIPTION
	Allow you to export a list of Drivers from your computer.
	It will list each service with the following informations: Device name, manufacturer, version, inf name
	Drivers list can be export to the following format: CSV, XLSX, XML, HTML

.EXAMPLE
PS Root\> Export-Drivers -Path C:\ -csv
The command above will export a Drivers list in CSV format in the folder C:\

.EXAMPLE
PS Root\> Export-Drivers -Path C:\ -xml
The command above will export a Drivers list in XML format in the folder C:\

.EXAMPLE
PS Root\> Export-Drivers -Path C:\ -html
The command above will export a Drivers list in HTML format in the folder C:\

.NOTES
    Author: Damien VAN ROBAEYS - @syst_and_deploy - http://www.systanddeploy.com
#>

Function Compare-Drivers
{
[CmdletBinding()]
Param(
	[Parameter(Mandatory=$true,ValueFromPipeline=$true, position=1)]
	[string] $Path,
	[Parameter(Mandatory=$true)]		
	[string] $File1,
	[Parameter(Mandatory=$true)]		
	[string] $File2,		
	[Switch] $XLS,
	[Switch] $HTML					
    )    
     Begin
    {		
		# If both files File1 and File2 are CSV
		If (($File1.contains("csv")) -and ($File2.contains("csv")))
			{
				$file1_content = import-csv $File1	 
				$file2_content = import-csv $File2							
			}
			
		# If both files File1 and File2 are XML			
		ElseIf (($File1.contains("xml")) -and ($File2.contains("xml")))
			{
				$file1_content = import-clixml $File1	 
				$file2_content = import-clixml $File2							
			}	
	
		# If File1 is XML and File2 is CSV			
		ElseIf (($File1.contains("xml")) -and ($File2.contains("csv")))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You have specified two different files to compare" -foregroundcolor "yellow"		
				write-host " !!! The first file is a XML and the second file is an CSV" -foregroundcolor "yellow"		 								
				write-host " !!! You need to specify two CSV or two XML files for the comparison" -foregroundcolor "yellow"										
				write-host "*******************************************************************"			
				break
			}
			
		# If File1 is CSV and File2 is XML						
		ElseIf (($File1.contains("csv")) -and ($File2.contains("xml")))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You have specified two different files to compare" -foregroundcolor "yellow"		
				write-host " !!! The first file is a CSV and the second file is an XML" -foregroundcolor "yellow"	
				write-host " !!! You need to specify two CSV or two XML files for the comparison" -foregroundcolor "yellow"										
				write-host "*******************************************************************"		
				break				
			}		
	
		# If there is no output format specified
		If ((-not $XLS) -and (-not $HTML))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You need to specific an output format. " -foregroundcolor "yellow"		
				write-host " !!! Use the switch -XLS to export in CSV or XLS format " -foregroundcolor "yellow"		 	
				write-host " !!! Use the switch -html to export in HTML format " -foregroundcolor "yellow"										
				write-host "*******************************************************************"					
			}		

	
		# If there output format is XLS					
		If($XLS)
			{			
				Try
					{
						# Check if an Excel process is already running, if yes it will keep the ID in the variable $excel_ID_Before
						$Excel_Process_Before = Get-Process Excel -ErrorAction SilentlyContinue
						$excel_ID_Before = $Excel_Process_Before.id						
					
						$Excel_Test = new-object -comobject excel.application
						$Excel_value = $True 
						write-host ""		
						write-host "********************************************************************************************************"	
						write-host "Drivers will be compared in XLS format" -foregroundcolor "Cyan"							
					}

				Catch 
					{
						$Excel_value = $false 					
					}		
			}	
			
		# If there output format is HTML			
		ElseIf($HTML)
			{
				write-host ""		
				write-host "********************************************************************************************************"	
				write-host "Drivers will be exported in HTML format" -foregroundcolor "Cyan"				
			}				
    }

	
    Process
    {
		If($XLS)
			{									
				$Compare_Drivers_ToXLS = $Path + "\" + "Compare_Drivers.xlsx"	

				$nbnewfile1 = 0
				$nbnewfile2 = 0		
				$nbdiffver = 0
				$nbsame = 0		
			
				$equal = compare-object $file1_content $file2_content -property devicename, driverversion -includeequal  | Where {$_.SideIndicator -eq "=="} | 
				   Group-Object -Property devicename | % { New-Object psobject -Property @{		
						Name=$_.group[0].devicename	
						Version=$_.group[0].driverversion			
					}}  | Select Name, Version | export-csv -encoding UTF8 -notype  $Temp_CSV_Equal 
									
				$nbsame = (compare-object $file1_content $file2_content -property devicename, driverversion -includeequal  | Where {$_.SideIndicator -eq "=="}|measure-object).count 							
					
				### NEW DRIVERS IN FILE 1
				$found = $false
				$newin1 = Foreach ($line1 in $file1_content)
					{
						$found = $false
						ForEach ($line2 in $file2_content)
							{
								IF ($line1.devicename -eq $line2.devicename)
									{
										$found = $true
										break
									}
							}

						IF (-not $found) 
							{
								New-Object -TypeName PSObject -Property @{
									Name = $line1.devicename
									Version = $line1.driverversion
									}
								$nbnewfile1 = $nbnewfile1 + 1
							}
					}	

				### NEW DRIVERS IN FILE 2
				$found = $false
				$newin2 = Foreach ($line2 in $file2_content)
					{
						$found = $false
						ForEach ($line1 in $file1_content)
							{
								IF ($line1.devicename -eq $line2.devicename)
									{
										$found = $true
										break
									}
							}

						IF (-not $found) 
							{
								New-Object -TypeName PSObject -Property @{
									Name = $line2.devicename
									Version = $line2.driverversion
									}
								$nbnewfile2 = $nbnewfile2 + 1	
							}
					}									
					

				### SAME DRIVERS BUT DIFFERENT VERSION
				$Diff_version = ForEach ($line1 in $file1_content)
				{
					ForEach ($line2 in $file2_content)
						{
							IF ($line1.devicename -eq $line2.devicename)
								{
									IF ($line1.driverversion -ne $line2.driverversion)
										{		
											New-Object -TypeName PSObject -Property @{
												Name = $line1.devicename
												Version_F1 = $line1.driverversion
												Version_F2 = $line2.driverversion 
												}  	
											$nbdiffver = $nbdiffver + 1	
										}
										Break
								}
						}												
				}	
		
				$newin1 | Select Name, Version | export-csv -encoding UTF8 -notype $Temp_CSV_MissingInFile2
				$newin2 | Select Name, Version | export-csv -encoding UTF8 -notype $Temp_CSV_MissingInFile1	
				
				$Diff_version | select Name, Version_F1, Version_F2 | export-csv -encoding UTF8 -notype  $Temp_CSV_DiffVersion
									
				$xl = new-object -comobject excel.application
				$xl.visible = $false
				$xl.DisplayAlerts=$False

				$Workbook1 = $xl.workbooks.open($Temp_CSV_Equal)
				$Workbook2 = $xl.workbooks.open($Temp_CSV_MissingInFile1)
				$Workbook3 = $xl.workbooks.open($Temp_CSV_MissingInFile2) 
				$Workbook4 = $xl.workbooks.open($Temp_CSV_DiffVersion) 

				$WorkBook0 = $xl.WorkBooks.add()

				$sh1_wborkbook0 = $WorkBook0.sheets.item(1) # first sheet in destination workbook
				$sheetToCopy1 = $Workbook1.sheets.item(1) # source sheet to copy
				$sheetToCopy1.copy($sh1_wborkbook0) # copy source sheet to destination workbook

				$sh2_wborkbook0 = $WorkBook0.sheets.item(2) # first sheet in destination workbook
				$sheetToCopy2 = $Workbook2.sheets.item(1) # source sheet to copy
				$sheetToCopy2.copy($sh2_wborkbook0) # copy source sheet to destination workbook

				$sh3_wborkbook0 = $WorkBook0.sheets.item(3) # first sheet in destination workbook
				$sheetToCopy3 = $Workbook3.sheets.item(1) # source sheet to copy
				$sheetToCopy3.copy($sh3_wborkbook0) # copy source sheet to destination workbook
				
				$sh4_wborkbook0 = $WorkBook0.sheets.item(4) # first sheet in destination workbook
				$sheetToCopy4 = $Workbook4.sheets.item(1) # source sheet to copy
				$sheetToCopy4.copy($sh4_wborkbook0) # copy source sheet to destination workbook		
				
				$equalboth = $WorkBook0.Worksheets.item(1)
				$missingin1 = $WorkBook0.Worksheets.item(2)
				$missingin2 = $WorkBook0.Worksheets.item(3)
				$diffvers = $WorkBook0.Worksheets.item(4)
				
				$equalboth.name = 'Same drivers and versions'
				$missingin1.name = 'New drivers in file 2'
				$missingin2.name = 'New drivers in file 1'
				$diffvers.name = 'Different version'
				
				$equalboth.columns.autofit() | out-null
				$missingin1.columns.autofit() | out-null
				$missingin2.columns.autofit() | out-null
				$diffvers.columns.autofit() | out-null
				
				$Table_Equal = $equalboth.ListObjects.add( 1,$equalboth.UsedRange,0,1)	
				$equalboth.ListObjects.Item($Table_Equal.Name).TableStyle="TableStyleMedium6"	

				$Table_Miss1 = $missingin1.ListObjects.add( 1,$missingin1.UsedRange,0,1)	
				$missingin1.ListObjects.Item($Table_Miss1.Name).TableStyle="TableStyleMedium3"	
				
				$Table_Miss2 = $missingin2.ListObjects.add( 1,$missingin2.UsedRange,0,1)	
				$missingin2.ListObjects.Item($Table_Miss2.Name).TableStyle="TableStyleMedium5"	
				
				$Table_DiffVer = $diffvers.ListObjects.add( 1,$diffvers.UsedRange,0,1)	
				$diffvers.ListObjects.Item($Table_DiffVer.Name).TableStyle="TableStyleMedium8"					
				
				$WorkBook0.SaveAs($Compare_Drivers_ToXLS,51)
				$WorkBook0.Saved = $True
				
				$WorkBook0.Close($false)
				$xl.Quit()		
				
				# Check if an Excel process is running with ID different of the $excel_ID_Before ID process
				# If yes it will store IDs in the variable $Excel_Process_After
				# Then all Process in $Excel_Process_After
				$Excel_Process_After = Get-Process Excel | where {$_.id -ne $excel_ID_Before}												
				Foreach ($Process_XL in $Excel_Process_After)						
					{
						stop-process $Process_XL.id	
					}								
			}		

		ElseIf($HTML)
			{				
				$file1_name = split-path $File1 -leaf -resolve	
				$file2_name = split-path $File2 -leaf -resolve	
			
				$Compare_Drivers_ToHTML = $Path + "\" + "Compare_Drivers.html"

				$nbnewfile1 = 0
				$nbnewfile2 = 0		
				$nbdiffver = 0
				$nbsame = 0	
			
				$Title = "<p><span class=Main_Title>Drivers comparison between $file1_name and $file2_name</span><br><span class=subtitle>This document has been updated on $date</span><br><span class=module_version>Module version: CompareComputer $Module_version</span></p><br><br>"			
				
				### SAME DRIVERS BUT DIFFERENT VERSION
				$Diff_Version_Title = "<p class=notequal_list>Different driver versions between $file1_name and $file2_name</p>"						
				$Diff_version = ForEach ($line1 in $file1_content)
				{
					ForEach ($line2 in $file2_content)
						{
							IF ($line1.devicename -eq $line2.devicename)
								{
									IF ($line1.driverversion -ne $line2.driverversion)
										{		
											New-Object -TypeName PSObject -Property @{
												Name = $line1.devicename
												Version_F1 = $line1.driverversion
												Version_F2 = $line2.driverversion 
												}  	
											$nbdiffver = $nbdiffver + 1												
										}
										Break
								}
						}												
				}

				### NEW DRIVERS IN FILE 1
				$Result_newIn1 = "<p class=New_object>New drivers in $file1_name</p>"						
				$found = $false
				$newin1 = Foreach ($line1 in $file1_content)
					{
						$found = $false
						ForEach ($line2 in $file2_content)
							{
								IF ($line1.devicename -eq $line2.devicename)
									{
										$found = $true
										break
									}
							}

						IF (-not $found) 
							{
								New-Object -TypeName PSObject -Property @{
									Name = $line1.devicename
									Version = $line1.driverversion
									}
								$nbnewfile1 = $nbnewfile1 + 1	
							}
					}	

				### NEW DRIVERS IN FILE 2
				$Result_newIn2 = "<p class=New_object>New drivers in $file2_name</p>"						
				$found = $false
				$newin2 = Foreach ($line2 in $file2_content)
					{
						$found = $false
						ForEach ($line1 in $file1_content)
							{
								IF ($line1.devicename -eq $line2.devicename)
									{
										$found = $true
										break
									}
							}

						IF (-not $found) 
							{
								New-Object -TypeName PSObject -Property @{
									Name = $line2.devicename
									Version = $line2.driverversion
									}
								$nbnewfile2 = $nbnewfile2 + 1
							}
					}		

				### SAME DRIVERS AND SAME VERSION									
					$Same_Values_Title = "<p class=equal_list>Same drivers and same version</p>"
					$Same_Values = compare-object $file1_content $file2_content -includeequal -property devicename, driverversion | Where {$_.SideIndicator -eq "=="} | 
					  Group-Object -Property devicename | % { New-Object psobject -Property @{
						Name=$_.group[0].devicename
						Version=$_.group[0].driverversion
						}}  | ConvertTo-HTML -Fragment

				$nbsame = (compare-object $file1_content $file2_content -property devicename, driverversion -includeequal  | Where {$_.SideIndicator -eq "=="}|measure-object).count 		

				$Resume_Table =	New-Object -TypeName PSObject -Property @{	
								"Same drivers and versions" = $nbsame
								"Drivers with different versions" = $nbdiffver 								
								"New drivers in file 1" = $nbnewfile1
								"New drivers in file 2" = $nbnewfile2 	
							}
		
				$Resume = $Resume_Table | Select "Same drivers and versions", "Drivers with different versions", "New drivers in file 1", "New drivers in file 2" | convertto-html -CSSUri $CSS_File
																	
				# Part to check what to display in the report
				
				# If there is no same drivers with same versions between both files, the part same drivers and same versions will be hidden
				If (($nbsame -eq 0)) 
					{
						$Same_Values_Title = ""
						$Same_Values = ""
					}				
				
				# If there are new drivers in both file 1 and 2, both parts new drivers in file 1 and new drivers in file 2 will be displayed
				If (($nbnewfile1 -ne 0) -and ($nbnewfile2 -ne 0)) 
					{
						$html3 = $newin1 | select  Name, Version | convertto-html -CSSUri $CSS_File							
						$html4 = $newin2 | select  Name, Version | convertto-html -CSSUri $CSS_File	
					}
			
				# If there is no new drivers in both file 1 and 2, both parts new drivers in file 1 and new drivers in file 2 will be hidden
				If (($nbnewfile1 -eq 0) -and ($nbnewfile2 -eq 0)) 
					{
						$Result_newIn1 = ""					
						$Result_newIn2 = ""					
					}
					
				# If there is no new driver in file 1 and there are new drivers in file 2, the part new drivers in file 1 will be hidden and the part new drivers in file 2 will be displayed 	
				If (($nbnewfile1 -eq 0) -and ($nbnewfile2 -ne 0)) 
					{
						$Result_newIn1 = ""	
						$html4 = $newin2 | select  Name, Version | convertto-html -CSSUri $CSS_File							
					}	
					
				# If there is no new driver in file 2 and there are new drivers in file 1, the part new drivers in file 2 will be hidden and the part new drivers in file 1 will be displayed 	
				If (($nbnewfile1 -ne 0) -and ($nbnewfile2 -eq 0)) 
					{					
						$Result_newIn2 = ""	
						$html3 = $newin1 | select  Name, Version | convertto-html -CSSUri $CSS_File																			
					}		
					
				# If there is no same driver with different versions between both files, the part same drivers with different versions will be hidden	
				If (($nbdiffver -eq 0)) 
					{
						$Diff_Version_Title = ""
					}

				# If there are same drivers with different versions between both files, the part same drivers with different versions will be displayed						
				If (($nbdiffver -ne 0)) 
					{
						$html1 = $Diff_version | select Name, Version_F1, Version_F2 | convertto-html -CSSUri $CSS_File
					}					
	
				$html_final = convertto-html -body "$Title<span class=Resume_Title>Resume values</span><br><br>$Resume <br><br>			
				<div id=left>$Same_Values_Title $Same_Values</div>				
				<div id=right_drivers>$Diff_Version_Title $html1 
					<br>
					$Result_newIn1 $html3
					<br>
					$Result_newIn2 $html4
				</div>									
				" -CSSUri $CSS_File
				$html_final | out-file -encoding ASCII $Compare_Drivers_ToHTML
				invoke-expression $Compare_Drivers_ToHTML					
			}				
	}	
	

   end
    {
		If($XLS)
			{			
				If ($Excel_value -eq $True)
					{											
						write-host "Drivers have been compared in XLS format" -foregroundcolor "Cyan"
						write-host "********************************************************************************************************"	
						If (($nbnewfile1 -ne 0) -or ($nbnewfile2 -ne 0) -or ($nbdiffstate -ne 0) -or ($nbdiffstartmode -ne 0))
							{
								write-host "Comparison Status: Some Drivers seems to be different between the two files !!!" -foregroundcolor "yellow"	 							
							}
						ElseIf (($nbnewfile1 -eq 0) -or ($nbnewfile2 -eq 0) -or ($nbdiffstate -eq 0) -or ($nbdiffstartmode -eq 0))
							{
								write-host "Comparison Status: All Drivers are similar !!!" -foregroundcolor "green"	 							
							}	
							
						write-host "********************************************************************************************************"	
						write-host "See below the results of the comparison" -foregroundcolor "Cyan"
						write-host ""							
						write-host "Drivers with equal values : $nbsame"
						write-host "New Drivers in file 1 : $nbnewfile1"
						write-host "New Drivers in file 2 : $nbnewfile2"
						write-host "Drivers with different version : $nbdiffver"		
						write-host "********************************************************************************************************"	
					}
				Else
					{
						write-host "Excel seems to be not installed"	
					}
			}					
			
		ElseIf($HTML)
			{										
				write-host "Drivers have been compared in HTML format" -foregroundcolor "Cyan"
				write-host "********************************************************************************************************"	
				If (($nbnewfile1 -ne 0) -or ($nbnewfile2 -ne 0) -or ($nbdiffver -ne 0))
					{
						write-host "Comparison Status: Some Drivers seems to be different between the two files !!!" -foregroundcolor "yellow"	 							
					}
				ElseIf (($nbnewfile1 -eq 0) -and ($nbnewfile2 -eq 0) -and ($nbdiffver -eq 0))
					{
						write-host "Comparison Status: All Drivers are similar !!!" -foregroundcolor "green"	 							
					}	
					
				write-host "********************************************************************************************************"	
				write-host "See below the results of the comparison" -foregroundcolor "Cyan"
				write-host ""							
				write-host "Drivers with equal values : $nbsame"
				write-host "New Drivers in file 1 : $nbnewfile1"
				write-host "New Drivers in file 2 : $nbnewfile2"
				write-host "Drivers with different version : $nbdiffver"		
				write-host "********************************************************************************************************"							
			}
	}	
	
}



















<#.Synopsis
	The Compare-Software function allows you to export a Software list from your computer. 
.DESCRIPTION
	Allow you to export a list of Software from your computer.
	It will list each service with the following informations: Name, Version
	Software list can be export to the following format: CSV, XLSX, XML, HTML

.EXAMPLE
PS Root\> Export-Software -Path C:\ -csv
The command above will export a Software list in CSV format in the folder C:\

.EXAMPLE
PS Root\> Export-Software -Path C:\ -xml
The command above will export a Software list in XML format in the folder C:\

.EXAMPLE
PS Root\> Export-Software -Path C:\ -html
The command above will export a Software list in HTML format in the folder C:\

.NOTES
    Author: Damien VAN ROBAEYS - @syst_and_deploy - http://www.systanddeploy.com
#>

Function Compare-Software
{
[CmdletBinding()]
Param(
	[Parameter(Mandatory=$true,ValueFromPipeline=$true, position=1)]
	[string] $Path,
	[Parameter(Mandatory=$true)]		
	[string] $File1,
	[Parameter(Mandatory=$true)]		
	[string] $File2,		
	[Switch] $XLS,
	[Switch] $HTML					
    )
    
     Begin
    {		
		# If both files File1 and File2 are CSV
		If (($File1.contains("csv")) -and ($File2.contains("csv")))
			{
				$file1_content = import-csv $File1	 
				$file2_content = import-csv $File2							
			}
			
		# If both files File1 and File2 are XML			
		ElseIf (($File1.contains("xml")) -and ($File2.contains("xml")))
			{
				$file1_content = import-clixml $File1	 
				$file2_content = import-clixml $File2							
			}	
	
		# If File1 is XML and File2 is CSV			
		ElseIf (($File1.contains("xml")) -and ($File2.contains("csv")))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You have specified two different files to compare" -foregroundcolor "yellow"		
				write-host " !!! The first file is a XML and the second file is an CSV" -foregroundcolor "yellow"		 								
				write-host " !!! You need to specify two CSV or two XML files for the comparison" -foregroundcolor "yellow"										
				write-host "*******************************************************************"			
				break
			}
			
		# If File1 is CSV and File2 is XML						
		ElseIf (($File1.contains("csv")) -and ($File2.contains("xml")))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You have specified two different files to compare" -foregroundcolor "yellow"		
				write-host " !!! The first file is a CSV and the second file is an XML" -foregroundcolor "yellow"	
				write-host " !!! You need to specify two CSV or two XML files for the comparison" -foregroundcolor "yellow"										
				write-host "*******************************************************************"		
				break				
			}		
	
		# If there is no output format specified
		If ((-not $XLS) -and (-not $HTML))
			{
				write-host ""		
				write-host "*******************************************************************"	
				write-host " !!! You need to specific an output format. " -foregroundcolor "yellow"		
				write-host " !!! Use the switch -XLS to export in CSV or XLS format " -foregroundcolor "yellow"		 	
				write-host " !!! Use the switch -html to export in HTML format " -foregroundcolor "yellow"										
				write-host "*******************************************************************"					
			}		
	
	
		# If there output format is XLS					
		If($XLS)
			{			
				Try
					{
						# Check if an Excel process is already running, if yes it will keep the ID in the variable $excel_ID_Before
						$Excel_Process_Before = Get-Process Excel -ErrorAction SilentlyContinue
						$excel_ID_Before = $Excel_Process_Before.id						
					
						$Excel_Test = new-object -comobject excel.application
						$Excel_value = $True 
						write-host ""		
						write-host "********************************************************************************************************"	
						write-host "Software will be compared in XLS format" -foregroundcolor "Cyan"						
					}

				Catch 
					{
						$Excel_value = $false 					
					}		
			}	
			
		# If there output format is HTML			
		ElseIf($HTML)
			{
				write-host ""		
				write-host "********************************************************************************************************"	
				write-host "Software will be compared in HTML format" -foregroundcolor "Cyan"				
			}				
    }

	
	
    Process
    {
		If($XLS)
			{
				If ($Excel_value -eq $True)												
					{						
						$file1_name = split-path $File1 -leaf -resolve	
						$file2_name = split-path $File2 -leaf -resolve	
								
						$Compare_Software_ToXLS = $Path + "\" + "Compare_Software.xlsx"	
						
						$nbnewfile1 = 0
						$nbnewfile2 = 0		
						$nbdiffvers = 0
						$nbsame = 0	
					
						$equal = compare-object $file1_content $file2_content -property DisplayName, displayversion -includeequal  | Where {$_.SideIndicator -eq "=="} | 
						   Group-Object -Property DisplayName | % { New-Object psobject -Property @{		
								Name=$_.group[0].DisplayName	
								Version=$_.group[0].displayversion		
							# }}  | export-csv -encoding UTF8 -notype $Temp_CSV_Equal 
							}}  | Select Name, Version | export-csv -encoding UTF8 -notype  $Temp_CSV_Equal 							
														
						$nbsame = (compare-object $file1_content $file2_content -property DisplayName, displayversion -includeequal  | Where {$_.SideIndicator -eq "=="}|measure-object).count 		
						
						### NEW SOFTWARE IN FILE 1
						$found = $false
						
						$newin1 = Foreach ($line1 in $file1_content)
							{
								$found = $false
								ForEach ($line2 in $file2_content)
									{
										IF ($line1.DisplayName -eq $line2.DisplayName)
											{
												$found = $true
												break
											}
									}

								IF (-not $found) 
									{
										New-Object -TypeName PSObject -Property @{
											Name = $line1.DisplayName
											Version = $line1.displayversion
											}
										$nbnewfile1 = $nbnewfile1 + 1							
									}					
							}
							
						### NEW SOFTWARE IN FILE 2
						$found = $false
						$newin2 = Foreach ($line2 in $file2_content)
							{
								$found = $false
								ForEach ($line1 in $file1_content)
									{
										IF ($line1.DisplayName -eq $line2.DisplayName)
											{
												$found = $true
												break
											}
									}

								IF (-not $found) 
									{
										New-Object -TypeName PSObject -Property @{
											Name = $line2.DisplayName
											Version = $line2.displayversion
											}
										$nbnewfile2 = $nbnewfile2 + 1									
									}
							}	
										
						### SAME SOFTWARE BUT DIFFERENT VERSION
						$Diff_version = ForEach ($line1 in $file1_content)
						{
							ForEach ($line2 in $file2_content)
								{
									IF ($line1.DisplayName -eq $line2.DisplayName)
										{
											IF ($line1.displayversion -ne $line2.displayversion)
												{		
													New-Object -TypeName PSObject -Property @{
														Name = $line1.DisplayName
														Version_F1 = $line1.displayversion
														Version_F2 = $line2.displayversion 
														}  
													$nbdiffver = $nbdiffver + 1
												}
												Break
										}
								}												
						}	
																		
						$Diff_version | Select Name, Version_F1, Version_F2 | export-csv -encoding UTF8 -notype $Temp_CSV_DiffVersion
						$newin1 | Select Name, Version | export-csv -encoding UTF8 -notype $Temp_CSV_MissingInFile2
						$newin2 | Select Name, Version | export-csv -encoding UTF8 -notype $Temp_CSV_MissingInFile1
																																											
						$xl = new-object -comobject excel.application
						$xl.visible = $false
						$xl.DisplayAlerts=$False

						$Workbook1 = $xl.workbooks.open($Temp_CSV_Equal)
						$Workbook2 = $xl.workbooks.open($Temp_CSV_MissingInFile1)
						$Workbook3 = $xl.workbooks.open($Temp_CSV_MissingInFile2) 
						$Workbook4 = $xl.workbooks.open($Temp_CSV_DiffVersion) 

						$WorkBook0 = $xl.WorkBooks.add()

						$sh1_wborkbook0 = $WorkBook0.sheets.item(1) # first sheet in destination workbook
						$sheetToCopy1 = $Workbook1.sheets.item(1) # source sheet to copy
						$sheetToCopy1.copy($sh1_wborkbook0) # copy source sheet to destination workbook

						$sh2_wborkbook0 = $WorkBook0.sheets.item(2) # first sheet in destination workbook
						$sheetToCopy2 = $Workbook2.sheets.item(1) # source sheet to copy
						$sheetToCopy2.copy($sh2_wborkbook0) # copy source sheet to destination workbook

						$sh3_wborkbook0 = $WorkBook0.sheets.item(3) # first sheet in destination workbook
						$sheetToCopy3 = $Workbook3.sheets.item(1) # source sheet to copy
						$sheetToCopy3.copy($sh3_wborkbook0) # copy source sheet to destination workbook
						
						$sh4_wborkbook0 = $WorkBook0.sheets.item(4) # first sheet in destination workbook
						$sheetToCopy4 = $Workbook4.sheets.item(1) # source sheet to copy
						$sheetToCopy4.copy($sh4_wborkbook0) # copy source sheet to destination workbook							

						$equalboth = $WorkBook0.Worksheets.item(1)
						$missingin1 = $WorkBook0.Worksheets.item(2)
						$missingin2 = $WorkBook0.Worksheets.item(3)
						$diffvers = $WorkBook0.Worksheets.item(4)
						
						$equalboth.name = 'Same software and versions'
						$missingin1.name = 'New software in file 2'
						$missingin2.name = 'Software missing in file 2'
						$diffvers.name = 'Different versions'

						$equalboth.columns.autofit() | out-null
						$missingin1.columns.autofit() | out-null
						$missingin2.columns.autofit() | out-null
						$diffvers.columns.autofit() | out-null
						
						$Table_Equal = $equalboth.ListObjects.add( 1,$equalboth.UsedRange,0,1)	
						$equalboth.ListObjects.Item($Table_Equal.Name).TableStyle="TableStyleMedium6"	

						$Table_Miss1 = $missingin1.ListObjects.add( 1,$missingin1.UsedRange,0,1)	
						$missingin1.ListObjects.Item($Table_Miss1.Name).TableStyle="TableStyleMedium3"	

						$Table_Miss2 = $missingin2.ListObjects.add( 1,$missingin2.UsedRange,0,1)	
						$missingin2.ListObjects.Item($Table_Miss2.Name).TableStyle="TableStyleMedium5"	
						
						$Table_DiffVer = $diffvers.ListObjects.add( 1,$diffvers.UsedRange,0,1)	
						$diffvers.ListObjects.Item($Table_DiffVer.Name).TableStyle="TableStyleMedium8"							

						$WorkBook0.SaveAs($Compare_Software_ToXLS,51)
						$WorkBook0.Saved = $True
						$xl.Quit()	

						# Check if an Excel process is running with ID different of the $excel_ID_Before ID process
						# If yes it will store IDs in the variable $Excel_Process_After
						# Then all Process in $Excel_Process_After
						$Excel_Process_After = Get-Process Excel | where {$_.id -ne $excel_ID_Before}												
						Foreach ($Process_XL in $Excel_Process_After)						
							{
								stop-process $Process_XL.id	
							}						
					}															
			}		

		ElseIf($HTML)
			{			
				$file1_name = split-path $File1 -leaf -resolve	
				$file2_name = split-path $File2 -leaf -resolve						
													
				$Compare_Software_ToHTML = $Path + "\" + "Compare_Software.html"			
								
				$nbnewfile1 = 0
				$nbnewfile2 = 0		
				$nbdiffver = 0
				$nbsame = 0
				
				$Title = "<p><span class=Main_Title>Software comparison between $file1_name and $file2_name</span><br><span class=subtitle>This document has been updated on $date</span><br><span class=module_version>Module version: CompareComputer $Module_version</span></p><br><br>"			
									
				### SAME SOFTWARE BUT DIFFERENT VERSION
				$Diff_Version_Title = "<p class=notequal_list>Different versions between both files</p>"						
				$Diff_version = ForEach ($line1 in $file1_content)
				{
					ForEach ($line2 in $file2_content)
						{
							IF ($line1.DisplayName -eq $line2.DisplayName)
								{
									IF ($line1.displayversion -ne $line2.displayversion)
										{		
											New-Object -TypeName PSObject -Property @{
												Name = $line1.DisplayName
												Version_F1 = $line1.displayversion
												Version_F2 = $line2.displayversion 
												}  	
											$nbdiffver = $nbdiffver + 1
										}
										Break
								}
						}												
				}

				### NEW SOFTWARE IN FILE 1
				$Result_newIn1 = "<p class=New_object>Software from file1 and missing in file2</p>"										
				$found = $false
				$newin1 = Foreach ($line1 in $file1_content)
					{
						$found = $false
						ForEach ($line2 in $file2_content)
							{
								IF ($line1.DisplayName -eq $line2.DisplayName)
									{
										$found = $true
										break
									}
							}

						IF (-not $found) 
							{
								New-Object -TypeName PSObject -Property @{
									Name = $line1.DisplayName
									Version = $line1.displayversion
									}
								$nbnewfile1 = $nbnewfile1 + 1	
							}
					}	

				### NEW SOFTWARE IN FILE 2
				$Result_newIn2 = "<p class=New_object>New software in $file2_name</p>"						
				$found = $false
				$newin2 = Foreach ($line2 in $file2_content)
					{
						$found = $false
						ForEach ($line1 in $file1_content)
							{
								IF ($line1.DisplayName -eq $line2.DisplayName)
									{
										$found = $true
										break
									}
							}

						IF (-not $found) 
							{
								New-Object -TypeName PSObject -Property @{
									Name = $line2.DisplayName
									Version = $line2.displayversion
									}
								$nbnewfile2 = $nbnewfile2 + 1								
							}
					}		

				### SAME SOFTWARE AND SAME VERSION									
					$Same_values_Title = "<p class=equal_list>Same softwares and versions</p>"
					$Same_values = compare-object $file1_content $file2_content -includeequal -property DisplayName, displayversion | Where {$_.SideIndicator -eq "=="} | 
					  Group-Object -Property DisplayName | % { New-Object psobject -Property @{
						Name=$_.group[0].DisplayName
						Version=$_.group[0].displayversion
						}}  | Select Name, Version | ConvertTo-HTML -Fragment

				$nbsame = (compare-object $file1_content $file2_content -property DisplayName, displayversion -includeequal  | Where {$_.SideIndicator -eq "=="}|measure-object).count 												

				$Resume_Table =	New-Object -TypeName PSObject -Property @{	
								"Same software and versions" = $nbsame
								"Software with different versions" = $nbdiffver 								
								"Software from file1 and missing in file 2" = $nbnewfile1
								"New software in file 2" = $nbnewfile2 								
							}
							
				$Resume = $Resume_Table | Select "Same software and versions", "Software with different versions", "Software from file1 and missing in file 2", "New software in file 2" | convertto-html -CSSUri $CSS_File
										
				# Part to check what to display in the report
				
				# If there is no same software with same version between both files, the part same software with same versions will be hidden
				If (($nbsame -eq 0)) 
					{
						$Same_Values_Title = ""
						$Same_Values = ""
					}					
				
				# If there are new software in both files 1 and 2, both parts new software in file 1 and new software in file 2 will be displayed
				If (($nbnewfile1 -ne 0) -and ($nbnewfile2 -ne 0)) 
					{
						$html3 = $newin1 | select  Name, Version | convertto-html -CSSUri $CSS_File							
						$html4 = $newin2 | select  Name, Version | convertto-html -CSSUri $CSS_File	
					}
			
				# If there is no new software in both files 1 and 2, both parts new software in file 1 and new software in file 2 will be hidden
				If (($nbnewfile1 -eq 0) -and ($nbnewfile2 -eq 0)) 
					{
						$Result_newIn1 = ""					
						$Result_newIn2 = ""					
					}
					
				# If there is no new software in file 1 and there are new software in file 2, the part new software in file 1 will be hidden and the part new software in file 2 will be displayed	
				If (($nbnewfile1 -eq 0) -and ($nbnewfile2 -ne 0)) 
					{
						$Result_newIn1 = ""	
						$html4 = $newin2 | select  Name, Version | convertto-html -CSSUri $CSS_File							
					}	
					
				# If there is no new software in file 2 and there are new software in file 1, the part new software in file 2 will be hidden and the part new software in file 1 will be displayed	
				If (($nbnewfile1 -ne 0) -and ($nbnewfile2 -eq 0)) 
					{					
						$Result_newIn2 = ""	
						$html3 = $newin1 | select  Name, Version | convertto-html -CSSUri $CSS_File																			
					}		
					
				# If there is no same software with same versions betwwen both files, the part same software with same versions will be hidden	
				If (($nbdiffver -eq 0)) 
					{
						$Diff_Version_Title = ""
					}

				# If there are same software with same versions betwwen both files, the part same software with same versions will be displayed						
				If (($nbdiffver -ne 0)) 
					{
						$html1 = $Diff_version | select Name, Version_F1, Version_F2 | convertto-html -CSSUri $CSS_File
					}					

				$html_final = convertto-html -body "$Title<span class=Resume_Title>Resume values</span><br><br>$Resume <br><br>			
				<div id=left_soft>$Same_Values_Title $Same_Values</div>				
				<div id=right_soft>$Diff_Version_Title $html1 
					<br>
					$Result_newIn1 $html3
					<br>
					$Result_newIn2 $html4
				</div>		
				" -CSSUri $CSS_File
				$html_final | out-file -encoding ASCII $Compare_Software_ToHTML
				invoke-expression $Compare_Software_ToHTML	
			}				
	}	

   end
    {
		If($XLS)
			{			
				If ($Excel_value -eq $True)
					{											
						write-host "Software have been compared in XLS format" -foregroundcolor "Cyan"
						write-host "********************************************************************************************************"	
						If (($nbnewfile1 -ne 0) -or ($nbnewfile2 -ne 0) -or ($nbdiffver -ne 0))
							{
								write-host "Comparison Status: Some Software seems to be different between the two files !!!" -foregroundcolor "yellow"	 							
							}
						ElseIf (($nbnewfile1 -eq 0) -and ($nbnewfile2 -eq 0) -and ($nbdiffver -eq 0))
							{
								write-host "Comparison Status: All Software are similar !!!" -foregroundcolor "green"	 							
							}	
							
						write-host "********************************************************************************************************"	
						write-host "See below the results of the comparison" -foregroundcolor "Cyan"
						write-host ""							
						write-host "Software with equal values : $nbsame"
						write-host "New Software in file 1 : $nbnewfile1"
						write-host "New Software in file 2 : $nbnewfile2"
						write-host "Same Software but with different version : $nbdiffver"		
						write-host "********************************************************************************************************"				
					}
				Else
					{
						write-host "Excel seems to be not installed"	
					}
					

			}					
			
		ElseIf($HTML)
			{										
				write-host "Software have been compared in HTML format" -foregroundcolor "Cyan"
				write-host "********************************************************************************************************"	
				If (($nbnewfile1 -ne 0) -or ($nbnewfile2 -ne 0) -or ($nbdiffver -ne 0))
					{
						write-host "Comparison Status: Some Software seems to be different between the two files !!!" -foregroundcolor "yellow"	 							
					}
				ElseIf (($nbnewfile1 -eq 0) -and ($nbnewfile2 -eq 0) -and ($nbdiffver -eq 0))
					{
						write-host "Comparison Status: All Software are similar !!!" -foregroundcolor "green"	 							
					}	
					
				write-host "********************************************************************************************************"	
				write-host "See below the results of the comparison" -foregroundcolor "Cyan"
				write-host ""							
				write-host "Software with equal values : $nbsame"
				write-host "New Software in file 1 : $nbnewfile1"
				write-host "New Software in file 2 : $nbnewfile2"
				write-host "Same Software but with different version : $nbdiffver"		
				write-host "********************************************************************************************************"						
			}
	}
	
}
