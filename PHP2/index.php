<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<title>Enrolment Display</title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	    <style type="text/css">
			<!--
			@media print{
				.box_header {
					padding: 1px;
					margin: 2px;
					font-size: 8pt;
					border-top: 1px solid #000000;
					border-right: 1px none #000000;
					border-bottom: 1px solid #000000;
					border-left: 1px solid #000000;
				}						
				.box {
					padding: 1px;
					margin: 2px;
					font-size: 8pt;
					border-top: 1px none #CCCCCC;
					border-right: 1px none #CCCCCC;
					border-bottom: 1px solid #CCCCCC;
					border-left: 1px solid #CCCCCC;
				}
			}
			@media screen{
				.box_header {
					padding: 1px;
					margin: 2px;
					font-size: x-small;
					border-top: 1px solid #000000;
					border-right: 1px none #000000;
					border-bottom: 1px solid #000000;
					border-left: 1px solid #000000;
					background-color: #CCCCCC;
				}						
				.box {
					padding: 1px;
					margin: 2px;
					font-size: x-small;
					border-top: 1px none #CCCCCC;
					border-right: 1px none #CCCCCC;
					border-bottom: 1px solid #CCCCCC;
					border-left: 1px solid #CCCCCC;
				}			
			}

			-->
        </style>
	</head>
	
	<body>
		<?PHP 
			include 'xmlprocess.php';

			function bracket_negative($figure) {
				if ($figure<0) {
					$figure = '('.(-$figure).')';
				}
				return $figure;
			}
			
			if ($year=='') {
				$year='0506';
			}
			
			$xml = file_get_contents("http://ptextend.ccb.ac.uk:83/gmpdev/EnrolmentDisplay/SelectCriteriaXML?year=$year");
			$my_records = new Records($xml);
		?>

		<form name="criteria" method="post" action="?app=<?php echo $app;?>">		
			<table>  
				<tr>
					<td class="hide_for_print" style="vertical-align:top;" >
						<table width="100%" border="0">
							<tr>
							  <td colspan="2" class="normal"><span class="menutitle">College Enrolments Summary</span></td>
						  	</tr>						
							<tr><td class="normal">&nbsp;</td></tr>
							<tr>														
								<td width="8%">PAM:</td> 
								<td width="17%">
									<select name="pam">
										<?php
											foreach ($my_records->rows as $row) {
												echo "<option value=\"$row[PERSON_CODE]\" ".($pam==$row[PERSON_CODE]?"selected":"").">$row[NAME]</option>";
											}
										?>
									</select>
								</td>
							</tr>
							<tr>
								<td>Year:</td>
								<td>
									<select name="year">
										<option value="0405" <?php echo ($year=="0405"?"selected":"")?>>2004/2005</option>
										<option value="0506" <?php echo ($year=="0506"?"selected":"")?>>2005/2006</option>
									</select>
								</td>
							</tr>
							<tr>
							  <td>Course Type:</td>
							  <td>
							  	<select name="course_type">
									<option value="0" <?php echo ($course_type==0?"selected":"")?>>All</option>
									<option value="1" <?php echo ($course_type==1?"selected":"")?>>Non GCSE</option>
									<option value="2" <?php echo ($course_type==2?"selected":"")?>>GCSE</option>
									<option value="3" <?php echo ($course_type==3?"selected":"")?>>Non EFL</option>
									<option value="4" <?php echo ($course_type==4?"selected":"")?>>EFL</option>
                              	</select>
							</td>
						  </tr>
							<tr>
							  <td>Course Attendance:</td>
							  <td>
								  <select name="attendance_type">
									<option value="0" <?php echo ($attendance_type==0?"selected":"")?>>All</option>
									<option value="1" <?php echo ($attendance_type==1?"selected":"")?>>Full Time</option>
									<option value="2" <?php echo ($attendance_type==2?"selected":"")?>>Part Time</option>
								  </select>
							  </td>
						  </tr>
							<tr>
							  <td>Additionalities</td>
							  <td>
								  <select name="additionality_type">
									<option value="0" <?php echo ($additionality_type==0?"selected":"")?>>No Additionalities</option>
									<option value="1" <?php echo ($additionality_type==1?"selected":"")?>>With Additionalities</option>
								  </select>
							  </td>
						  </tr>
							<tr>
							  <td>Funding Type:</td>
							  <td>
								  <select name="funding_type">
									<option value="0" <?php echo ($funding_type==0?"selected":"")?>>All</option>
									<option value="1" <?php echo ($funding_type==1?"selected":"")?>>LSC</option>
									<option value="2" <?php echo ($funding_type==2?"selected":"")?>>HE</option>
								  </select>
							  </td>
						  </tr>
							<tr>
							  <td>&nbsp;</td>
							  <td><input type="submit" name="Submit" value="Show Enrolments"></td>
						  </tr>
						</table>
					</td>
				
					<td style="vertical-align: top">
						<?php
							$years = array('0405' => '2004/2005', '0506' => '2005/2006');
		
							echo "<table width=\"100%\"  border=\"0\" cellpadding=\"0\" cellspacing=\"0\" >";
							echo "	<caption class=\"hide_for_screen\" align=\"left\" style=\"font-weight:bold;\">";
							echo "College Enrolments for $row[NAME] ".$years[$year];			
							echo "	</caption>";

							echo "<thead style=\"display: table-header-group;\">";
							echo "	<tr bordercolor=\"#000000\" bgcolor=\"#CCCCCC\" class=\"box\">";
							echo "		<th class=\"box_header\" width=\"5%\"><div align=\"center\">Centre</div></th>";
							echo "		<th class=\"box_header\" width=\"12%\"><div align=\"left\">Course Code</div> </th>";
							echo "		<th class=\"box_header\" width=\"40%\"><div align=\"left\">Course Title</div> </th>";
							echo "		<th class=\"box_header\" width=\"8%\"><div align=\"center\">Target Places</div> </th>";
							echo "		<th class=\"box_header\" width=\"8%\"><div align=\"center\">Total Places</div> </th>";
							echo "		<th class=\"box_header\" width=\"8%%\"><div align=\"center\">Students Enrolled</div> </th>";
							echo "		<th class=\"box_header\" width=\"8%\"><div align=\"center\">Target Shortfall</div></th>";
							echo "		<th class=\"box_header\" width=\"8%\"><div align=\"center\">Maximum Shortfall</div> </th>";
							echo "		<th class=\"box_header\" width=\"9%\"><div align=\"center\">Enrolled 16-18</div> </th>";
							echo "		<th class=\"box_header\" width=\"9%\"><div align=\"center\">Enrolled 19+</div> </th>";
							echo "		<th class=\"box_header\" width=\"9%\" style=\"border-right: 1px solid #000000;\"><div align=\"center\">Enrolled Overseas</div></th>";
							echo "	</tr>";
							echo "</thead>";

							echo "<tbody>";

							$xml = file_get_contents("http://ptextend.ccb.ac.uk:83/gmpdev/EnrolmentDisplay/CourseEnrolmentsXML?pam=$pam&year=$year&course_type=$course_type&attendance_type=$attendance_type&additionality_type=$additionality_type&funding_type=$funding_type");

							$my_records_courses = new Records($xml);				
							$row_index=0;
							$rows_per_page=5;
							$display_this_row=0;
							$page=0;
							
							$totals = array('TARGET_PLACES' => 0, 'TOTAL_PLACES' => 0, 'STUDENTS_ENROLLED' => 0, 'TARGET_SHORTFALL' => 0, 'MAXIMUM_SHORTFALL' => 0, 'ENROLLED_16_TO_18' => 0, 'ENROLED_OVER_19' => 0, 'ENROLLED_OVERSEAS' => 0);

							foreach ($my_records_courses->rows as $row) {
								if ($page!='') {
									if ($row_index>=($page*$rows_per_page)) {
										$display_this_row=true;
									} else {
										$display_this_row=false;
									}
								} else {
									$display_this_row=true;
								}
								
								if ($display_this_row) {								
									echo "<tr class=\"box\">";
									echo "    <td class=\"box\" style=\"text-align: center;\">".$my_records_courses->show_blank($row[CENTRECODE])."</td>";
									echo "    <td class=\"box\" style=\"text-align: left;\">".$my_records_courses->show_blank($row[COURSE_CODE])."</td>";
									echo "    <td class=\"box\" style=\"text-align: left;\">".$my_records_courses->show_blank($row[COURSE_TITLE])."</td>";
									echo "    <td class=\"box\" style=\"text-align: right;\">".$my_records_courses->show_blank($row[TARGET_PLACES])."</td>";
									echo "    <td class=\"box\" style=\"text-align: right;\">".$my_records_courses->show_blank($row[TOTAL_PLACES])."</td>";
									echo "    <td class=\"box\" style=\"text-align: right;\">".$my_records_courses->show_blank($row[STUDENTS_ENROLLED])."</td>";
									echo "    <td class=\"box\" style=\"text-align: right; background-color: #CCCCCC\">".bracket_negative($my_records_courses->show_blank($row[TARGET_SHORTFALL]))."</td>";
									echo "    <td class=\"box\" style=\"text-align: right;\">".bracket_negative($my_records_courses->show_blank($row[MAXIMUM_SHORTFALL]))."</td>";
									echo "    <td class=\"box\" style=\"text-align: right;\">".$my_records_courses->show_blank($row[ENROLLED_16_TO_18])."</td>";
									echo "    <td class=\"box\" style=\"text-align: right;\">".$my_records_courses->show_blank($row[ENROLLED_OVER_19])."</td>";
									echo "    <td class=\"box\" style=\"text-align: right; border-right: 1px solid #CCCCCC;\">".$my_records_courses->show_blank($row[ENROLLED_OVERSEAS])."</td>";
									echo "</tr>";
									
									$totals[TARGET_PLACES]+=$row[TARGET_PLACES];
									$totals[TOTAL_PLACES]+=$row[TOTAL_PLACES];
									$totals[STUDENTS_ENROLLED]+=$row[STUDENTS_ENROLLED];
									$totals[TARGET_SHORTFALL]+=$row[TARGET_SHORTFALL];
									$totals[MAXIMUM_SHORTFALL]+=$row[MAXIMUM_SHORTFALL];
									$totals[ENROLLED_16_TO_18]+=$row[ENROLLED_16_TO_18];
									$totals[ENROLLED_OVER_19]+=$row[ENROLLED_OVER_19];
									$totals[ENROLLED_OVERSEAS]+=$row[ENROLLED_OVERSEAS];
								}
								$row_index++;
							}
							
							// Totals
							echo "<tr class=\"box\">";
							echo "    <td></td>";
							echo "    <td></td>";
							echo "    <td class=\"box_header\" style=\"text-align: left; background-color: #CCCCCC;\"><div align='right'>TOTALS</div></td>";
							echo "    <td class=\"box_header\" style=\"text-align: right; background-color: #CCCCCC;\">".$my_records_courses->show_blank($totals[TARGET_PLACES])."</td>";
							echo "    <td class=\"box_header\" style=\"text-align: right; background-color: #CCCCCC;\">".$my_records_courses->show_blank($totals[TOTAL_PLACES])."</td>";
							echo "    <td class=\"box_header\" style=\"text-align: right; background-color: #CCCCCC;\">".$my_records_courses->show_blank($totals[STUDENTS_ENROLLED])."</td>";
							echo "    <td class=\"box_header\" style=\"text-align: right; background-color: #CCCCCC;\">".bracket_negative($my_records_courses->show_blank($totals[TARGET_SHORTFALL]))."</td>";
							echo "    <td class=\"box_header\" style=\"text-align: right; background-color: #CCCCCC;\">".bracket_negative($my_records_courses->show_blank($totals[MAXIMUM_SHORTFALL]))."</td>";
							echo "    <td class=\"box_header\" style=\"text-align: right; background-color: #CCCCCC;\">".$my_records_courses->show_blank($totals[ENROLLED_16_TO_18])."</td>";
							echo "    <td class=\"box_header\" style=\"text-align: right; background-color: #CCCCCC;\">".$my_records_courses->show_blank($totals[ENROLLED_OVER_19])."</td>";
							echo "    <td class=\"box_header\" style=\"text-align: right; background-color: #CCCCCC; border-right: 1px solid #CCCCCC;\">".$my_records_courses->show_blank($totals[ENROLLED_OVERSEAS])."</td>";
							echo "</tr>";
						
							echo "</tbody>";
							echo "</table>";

						?>
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>
