<?php
	header("Content-Type: text/html; charset=UTF-8");
	
	$xml_request = 
        "<?xml version='1.0' encoding='UTF-8' ?> 
        <QuoteRequest> 
                <QuoteID>373855313988635433</QuoteID> 
                <RequestingUserID>95543</RequestingUserID> 
                <PartnerUserID>1</PartnerUserID> 
                <Country>GB</Country> 
                <Project>SilverFox</Project> 
                <Quantity>3</Quantity> 
                <BoardLength>101</BoardLength> 
                <BoardWidth>102</BoardWidth> 
                <CuLayers>4</CuLayers> 
                <DSID>0</DSID> 
                <DlyRequired>10</DlyRequired> 
                <PcbMaterial>FR4</PcbMaterial> 
                <PcbThickness>1.6</PcbThickness> 
                <CuWeight>35</CuWeight> 
                <SolResTop>1</SolResTop> 
                <SolResBottom>1</SolResBottom> 
                <SolResColTop>green</SolResColTop> 
                <SolResColBottom>green</SolResColBottom> 
                <SilkScrTop>1</SilkScrTop> 
                <SilkScrBottom>1</SilkScrBottom> 
                <SilkScrColTop>white</SilkScrColTop> 
                <SilkScrColBottom>white</SilkScrColBottom> 
                <ElectricalTest>0</ElectricalTest> 
        </QuoteRequest>";

// http://www.pcboards.eu/app_custom/php/ds_curl.php
?>
<html>
	<body>
		<form action="http://www.wedirekt.de/FE_SVNSQUVMLUxFVkVSLTc3Ny1LUkFGVA/index.php/gateways/designspark/pcb_quote/" method="post">
			<input type="hidden" name="QuoteRequest" value="<?php echo $xml_request;?>">
			<input type="submit">
		</form>
		<?php
			echo htmlentities($xml_request);
		?>
	</body>
</html>