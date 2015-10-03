<?php
	$xml_request = "<?xml version='1.0' encoding='UTF-8' ?> <Request><QuoteID>515930953900597777</QuoteID><OrderValue>512.25</OrderValue></Request>";
?>
<html>
	<body>
		<form action="http://164.177.150.201/pcb_quote/report_log/" method="get">
			<input type="hidden" name="Request" value="<?php echo ($xml_request);?>">
			<input type="submit">
		</form>
	</body>
</html>

<?php echo htmlentities($xml_request);?>