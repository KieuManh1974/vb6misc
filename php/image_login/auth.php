<?php
	$check = array(0,1,2,3);

	$values = $_POST['submit'];
	sort($values, SORT_NUMERIC);


	$ok = true;
	foreach ($check as $index=>$item) {
		//echo $values[$index].'-'.$item.'<br>';
		if ($values[$index]!=$item) {
			$ok = false;
			break;
		}
	}

	if (!$ok) {
		echo "Create Account";
	} else {
		echo "You have logged in";
	}
?>