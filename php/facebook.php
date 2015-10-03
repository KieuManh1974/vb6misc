<?php
// http://stackoverflow.com/questions/4006289/facebook-api-how-to-update-the-status

	$email = $_GET['email'];
	$password = $_GET['password'];
	$status = urlencode( $_GET['status'] );

	$url = "http://graph.facebook.com/PROFILE_ID/feed";

	$session = curl_init();
	curl_setopt ( $session, CURLOPT_URL, $url );
	curl_setopt ( $session, CURLOPT_HTTPAUTH, CURLAUTH_BASIC );
	curl_setopt ( $session, CURLOPT_HEADER, false );
	curl_setopt ( $session, CURLOPT_USERPWD, $email . ":" . $password );
	curl_setopt ( $session, CURLOPT_RETURNTRANSFER, 1 );
	curl_setopt ( $session, CURLOPT_POST, 1);
	curl_setopt ( $session, CURLOPT_POSTFIELDS,"message=" . $status);
	$result = curl_exec ( $session );
	curl_close( $session );

	echo( $result );

?>

<?php

// http://www.9lessons.info/2011/01/facebook-graph-api-connect-with-php-and.html

	include('db.php');
	if($_SERVER["REQUEST_METHOD"] == "POST")
	{
	$status=$_POST['status'];
	$sql=mysql_query("select facebook_id,facebook_access_token from users where username='$user_session'");
	$row=mysql_fetch_array($sql);
	$facebook_id=$row['facebook_id'];
	$facebook_access_token=$row['facebook_access_token'];
	//Facebook Wall Update
	params = array('access_token'=>$facebook_access_token, 'message'=>$status);
	$url = "https://graph.facebook.com/$facebook_id/feed";
	$ch = curl_init();
	curl_setopt_array($ch, array(
	CURLOPT_URL => $url,
	CURLOPT_POSTFIELDS => $params,
	CURLOPT_RETURNTRANSFER => true,
	CURLOPT_SSL_VERIFYPEER => false,
	CURLOPT_VERBOSE => true
	));
	$result = curl_exec($ch);
	// End
	}
?>
//HTML
<form method="post" action="">
<textarea name="status"></textarea>
<input type="submit" value=" Update "/>
</form&gt;