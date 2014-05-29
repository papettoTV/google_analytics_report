<?php

ini_set( 'display_errors', 1 );
error_reporting(E_ALL);

require_once dirname(__FILE__) . "/Google_Analytics_Report.php";
$gar = new Google_Analytics_Report();

$gar->setAccounts();

$html ="<h1>Profile list</h1>";
$html .= "<ul>";
foreach($gar->getAccounts() as $ac){
	$html .= "<li>".$ac->name."(id:".$ac->getId().")<a href='dl_sample.php?id=".$ac->getId()."'>(Download report at last month)</a></li>";
}
$html .= "</ul>";
echo $html;

