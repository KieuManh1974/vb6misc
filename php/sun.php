<?php

	function closest($time) {
		$hour = (int) date("H",$time);
		$minute = (int) date("i",$time);
		$second = (int) date("s",$time);

		if ($second>=30) {
			$minute++;
		}
		if ($minute>=60) {
			$hour++;
			$minute=0;
		}

		return str_pad($hour, 2, "0", STR_PAD_LEFT).str_pad($minute, 2, "0", STR_PAD_LEFT);
	}

	function nearest_ten($time) {
		$hour = (int) date("H",$time);
		$minute = (int) date("i",$time);
		$second = (int) date("s",$time);

		$unit = $minute%10;
		if ($unit<5) {
			$minute_ten = $minute-$unit;
		} else {
			$minute_ten = $minute+(10-$unit);
		}

		if ($minute_ten>=60) {
			$hour++;
			$minute_ten-=60;
		}	

		return str_pad($hour, 2, "0", STR_PAD_LEFT).':'.str_pad($minute_ten, 2, "0", STR_PAD_LEFT).':00';
	}


	$sun_info = date_sun_info(time(), 50.824, -0.14028);

	// foreach ($sun_info as $title=>$info) {
	// 	echo $title.':'.date("Hi s",$info).'<br>';
	// }

	$lengths = array(31,28,31,30,31,30,31,31,30,31,30,31);
	
	$dates = array();
	for($month=1;$month<=12;$month++) {
		for($day=1;$day<=$lengths[$month-1];$day++) {
			$date = strtotime("2013-$month-$day");
			
			$sun_info = date_sun_info($date, 50.824, -0.14028);

			$diff = abs(strtotime(nearest_ten($sun_info['civil_twilight_end']))-strtotime(date("H:i:s",$sun_info['civil_twilight_end'])));

			$dates[] = array(date("d-M-Y",$date),nearest_ten($sun_info['civil_twilight_end']),$diff);
		}
	}

	echo '<br>';

	$diff=500;
	$first = true;
	foreach ($dates as $index=>$date) {
		if ($date[2]<=$diff) {
			$diff = $date[2];
			$first = true;
		} else {
			$diff=$date[2];
			if ($first) {
				//echo $dates[$index-1][0].' '.$dates[$index-1][1].'<br>';
				echo $dates[$index-1][1].'<br>';
			}
			$first = false;
		}
	}

// 50N49.440 0W8.417 Makemedia



?>