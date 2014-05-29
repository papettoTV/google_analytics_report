<?php

ini_set( 'display_errors', 1 );
error_reporting(E_ALL);

$dir = dirname(__FILE__);
require_once $dir . "/Google_Analytics_Report.php";
require_once $dir ."/Exel_Format_Report_Sample.php";

$gar = new Google_Analytics_Report();
$exel_report = new Exel_Format_Report();

$profileId = $_GET["id"];

// 先月
$analyze_date = mktime(0,0,0, date("m")-1, 1, date("Y"));
$analyze_year = date('Y',$analyze_date);
$analyze_month = date('m',$analyze_date);

// プロフィール設定
$gar->setProfileId($profileId);

// 期間設定
$gar->setTerm($analyze_year,$analyze_month);


// ****
// レポート作成開始
// ****
$exel_report->make($gar);

// エクセル出力
$exel_report->out();
exit;
