<?php

	$email = $_GET['email'];
	$password = $_GET['password'];
	$status = urlencode( $_GET['status'] );

	$url = "http://8.7.217.31/statuses/update.xml";

	$session = curl_init();
	curl_setopt ( $session, CURLOPT_URL, $url );
	curl_setopt ( $session, CURLOPT_HTTPAUTH, CURLAUTH_BASIC );
	curl_setopt ( $session, CURLOPT_HEADER, false );
	curl_setopt ( $session, CURLOPT_USERPWD, $email . ":" . $password );
	curl_setopt ( $session, CURLOPT_RETURNTRANSFER, 1 );
	curl_setopt ( $session, CURLOPT_POST, 1);
	curl_setopt ( $session, CURLOPT_POSTFIELDS,"status=" . $status);
	$result = curl_exec ( $session );
	curl_close( $session );

	echo( $result );

?>