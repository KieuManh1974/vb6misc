<?php
	include 'parser_compiler.php';

	$output = '';
	$input = '';
	$ok = true;
	if (isset($_POST['input'])) {
		$input = $_POST['input'];

		$p = new Parser();

		$rules_text = <<<'RULESTEXT'
			ws omit list set \32\\13\\10\ | min 0 | |
			identifier list set case abcdefghijklmnopqrstuvwxyz | | |	
			block and ws identifier ws { ws blocks ws } | |
			blocks list block min 0 | |
RULESTEXT;

		$p->CreateParser($rules_text);
		
		$s = new Stream($input);
		$result = $p->rules["block"]->Parse($s);

		$output = $result->text($s);
		$ok = $result->ok;
	}
?>
<!DOCTYPE html>
<html>
	<head>
		<meta charset='utf-8'> 
		<style>
			#output {
				background-color: #EEE;
				width:1000px;
				border: 1px solid #A0A0A0;
				margin-top:20px;
				min-height: 30px;
			}

			.red {
				color:red;
			}
		</style>
	</head>
	<body>
		<form action="enter.php" method="post">
			<textarea id="input" name="input" cols="150" rows="30"><?=$input;?></textarea>
			<input type="submit">
			<div id="output" class="<?=$ok?'':'red';?>">
				<pre><?= $output;?></pre>
			</div>
		</form>
	</body>
</html>