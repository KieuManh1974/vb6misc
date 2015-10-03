<?php

	$mysqli = new mysqli("164.177.132.96","s_bradley","mmsbd010212","rs_product_new");

	$offset = 0;
	$limit = 20;

	$dict = array();

for ($index=1; $index<10000; $index++) {
	$result = $mysqli->query("SELECT mpn FROM article LIMIT $limit OFFSET $offset");

	for ($row=0; $row<$result->num_rows; $row++) {
		$result->data_seek($row);
		$row_data = $result->fetch_assoc();
		$mpn = $row_data['mpn'];

		$mpn_length = strlen($mpn);
		for ($pos=0; $pos<($mpn_length-1); $pos++) {
			$word = substr($mpn,$pos,2);			
			//echo $word.'<br>';

			if (array_key_exists($word, $dict)) {
				$dict[$word][]=$offset;
			} else {
				$dict[$word]=array($offset);
			}

		}
	}

	$offset += $limit;

}


$pretty = function($v='',$c="&nbsp;&nbsp;&nbsp;&nbsp;",$in=-1,$k=null)use(&$pretty){$r='';if(in_array(gettype($v),array('object','array'))){$r.=($in!=-1?str_repeat($c,$in):'').(is_null($k)?'':"$k: ").'<br>';foreach($v as $sk=>$vl){$r.=$pretty($vl,$c,$in+1,$sk).'<br>';}}else{$r.=($in!=-1?str_repeat($c,$in):'').(is_null($k)?'':"$k: ").(is_null($v)?'&lt;NULL&gt;':"<strong>$v</strong>");}return$r;};

echo $pretty(array_keys($dict));


?>