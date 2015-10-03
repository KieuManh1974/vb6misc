<?php

	// $mysqli = new mysqli("164.177.132.96","s_bradley","mmsbd010212","rs_product_new");
	$mysqli = new mysqli("164.177.132.96","s_bradley","mmsbd010212","croogo14");

	$results = $mysqli->query("SELECT id, username, image FROM users WHERE image != '' ");

	$base_path_source = 'http://c1330478.r78.cf3.rackcdn.com/';
	$base_path_target = 'C:\\Development\\wamp\\tmp\\';

	$total_rows = $results->num_rows;
	// $total_rows = 100;

	for ($index=0; $index<$total_rows; $index++) {
		$results->data_seek($index);
		$row_data = $results->fetch_assoc();
		$username = $row_data['username'];
		$image = $row_data['image'];

		$ext = strtolower(substr($image, strlen($image)-3, 3));
		
		if (substr($image, 0, 6)=='files/') {
			$image = str_replace('files/ds/','',$image);
			$image = str_replace('files/all/','',$image);

			$path_source = $base_path_source.$image;
			$path_target = $base_path_target.$username.'.'.$ext;

			@copy($path_source, $path_target);
		}
	}


?>