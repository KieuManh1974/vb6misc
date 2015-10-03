<?php
	header("Content-Type: text/html; charset=UTF-8");
	
	$xml = 
	'<?xml version="1.0" encoding="UTF-8"?>
	<QuoteRequest>
	    <QuoteID>138991109282411267</QuoteID>
	    <RequestingUserID>1</RequestingUserID>
	    <PartnerUserID>4</PartnerUserID>
	    <Country>gb</Country>
	    <Project>ProjectBlackstone</Project>
	    <Quantity>2</Quantity>
	    <BoardLength>123</BoardLength>
	    <BoardWidth>210</BoardWidth>
	    <CuLayers>2</CuLayers>
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
	</QuoteRequest>';


	$ch = curl_init();
	curl_setopt($ch, CURLOPT_HEADER, 0);
	curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
	curl_setopt($ch, CURLOPT_URL, "http://pcbtrain.headserv.net/api/request_price/");
	curl_setopt($ch, CURLOPT_POST, 1);
	curl_setopt($ch, CURLOPT_POSTFIELDS, "QuoteRequest=".$xml);

    curl_setopt($ch, CURLOPT_FOLLOWLOCATION, 0);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);

	$response = curl_exec($ch);
	if ($response == false) {
	    throw new Exception("Failure response from curl");
	}
	$response = trim($response);

	echo $response;

?>