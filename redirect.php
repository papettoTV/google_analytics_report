<?php
session_start();

require_once dirname(__FILE__) . "/Google_Analytics_Report.php";

$client = new Google_Client();

$client->setClientId(GAR_API_CLIENT_ID);
$client->setClientSecret(GAR_API_CLIENT_SECRET);
$client->setDeveloperKey(GAR_API_DEVELOPER_KEY);
$client->setRedirectUri(GAR_API_REDIRECT_URI);
$client->setScopes(array('https://www.googleapis.com/auth/analytics'));

$root_dir = "http://".$_SERVER["HTTP_HOST"].dirname(parse_url($_SERVER["REQUEST_URI"], PHP_URL_PATH));

if (isset($_GET['code'])) {

	$_SESSION['google_analytics_report_token']=$client->authenticate($_GET['code']);

	$redirect = $root_dir."/accounts_list.php";
	header('Location: ' . filter_var($redirect, FILTER_SANITIZE_URL));
	exit;

}elseif (isset($_SESSION['google_analytics_report_token'])) {

	// logout
	unset($_SESSION['google_analytics_report_token']);
	$redirect = $root_dir."/accounts_list.php";
	header('Location: ' . filter_var($redirect, FILTER_SANITIZE_URL));
	exit;
}


$authUrl = $client->createAuthUrl();
print "<a class='login' href='$authUrl'>Connect Me!</a>";
exit;

