<?php
require_once dirname(__FILE__) . "/Google_Analytics_Report_Core.php";

class Google_Analytics_Report extends Google_Analytics_Report_Core{

	function __construct(){
		parent::__construct();
	}

	/*
	 * 日別レポート
	 */
	function getResultsDays(){
		return $this->getResultsCommon(
			'ga:pageviews,ga:visits,ga:goalCompletionsAll',
			"ga:day",
			"ga:day",
			31
		);
	}

	/*
	 * 曜日別レポート
	 */
	function getResultsWeek(){
		return $this->getResultsCommon(
			'ga:visits,ga:pageviews',
			"ga:dayOfWeek,ga:dayOfWeekName",
			"ga:dayOfWeek",
			7
		);
	}

	/*
	 * 時間別レポート
	 */
	function getResultsTime() {
		return $this->getResultsCommon(
			'ga:pageviews,ga:visits',
			"ga:hour",
			"ga:hour",
			24
		);
	}

	/*
	 * 参照元レポート(標準偏差付)
	 */
	function getResultsReferral($filter=""){
		$res = $this->getResultsCommon(
			'ga:visits,ga:pageviews,ga:avgSessionDuration,ga:percentNewVisits,ga:goalCompletionsAll,ga:newVisits,ga:bounces,ga:sessionDuration,ga:entrances',
			"ga:source,ga:medium",
			"-ga:visits",
			"",
			$filter
		);

		//獲得ポイント=セッション数*新規セッション率 + 総コンバージョン*10
		foreach($res["rows"] as $key => $v){
			$res["rows"][$key]["point"] = $v[1] * $v[4]/100 + $v[5]*10;
			$point[] = $res["rows"][$key]["point"];
		}
		// 標準偏差
		$standard_deviation = $this->statsStandardDeviation($point);
		// 平均
		$avg = array_sum($point) / count($point);

		foreach($res["rows"] as $key => $v){
			//キーワード有効偏差値 = (獲得ポイント - 平均ポイント)*10/ポイント標準偏差 + 50
			$res["rows"][$key]["deviation"] = ($v["point"] - $avg)*10/$standard_deviation + 50;
		}

		return $res;
	}

	/*
	 * キーワード別レポート（標準偏差付き）
	 */
	function getResultsKeyword(){
		$res = $this->getResultsCommon(
			'ga:pageviews,ga:visits,ga:avgSessionDuration,ga:percentNewVisits,ga:goalCompletionsAll,ga:goalConversionRateAll,ga:entrances,ga:bounces,ga:newVisits,ga:sessionDuration',
			"ga:Keyword",
			"-ga:visits",
			""
			,'ga:medium==organic'
		);

		//獲得ポイント=セッション数*新規セッション率 + 総コンバージョン*10
		foreach($res["rows"] as $key => $v){
			$res["rows"][$key]["point"] = $v[2] * $v[4]/100 + $v[5]*10;
			$point[] = $res["rows"][$key]["point"];
		}
		// 標準偏差
		$standard_deviation = $this->statsStandardDeviation($point);
		// 平均
		$avg = array_sum($point) / count($point);

		foreach($res["rows"] as $key => $v){
			//キーワード有効偏差値 = (獲得ポイント - 平均ポイント)*10/ポイント標準偏差 + 50
			$res["rows"][$key]["deviation"] = ($v["point"] - $avg)*10/$standard_deviation + 50;
		}

		return $res;
		
	}

	/*
	 * 閲覧開始ページ
	 */
	function getResultsEntrancePage($path=""){

		$filter="";
		if($path!=""){
			$filter="ga:pagePath==".$path;
		}
		// ga:exitRate はwebの数値とややずれありなので利用注意
		return $this->getResultsCommon(
			'ga:pageviews,ga:visits,ga:entrances,ga:bounces,ga:exitRate',
			"ga:landingPagePath,ga:pageTitle",
			"-ga:entrances"
			,""
			,$filter
		);
	}

	/*
	 * 離脱ページ
	 */
	function getResultsExitPage(){
		return $this->getResultsCommon(
			'ga:pageviews,ga:exits,ga:exitRate',
			"ga:pagePath,ga:pageTitle",
			"-ga:exits"
		);
	}

	/*
	 * 直帰ページ
	 */
	function getResultsBouncePage($path=""){

		$filter="";
		if($path!=""){
			$filter="ga:pagePath==".$path;
		}

		return $this->getResultsCommon(
			'ga:entrances,ga:bounces',
			"ga:pagePath,ga:pageTitle",
			"-ga:entrances",
			"",	
			$filter
		);
	}

	/*
	 * ページビュー数（標準偏差付き）
	 */
	function getResultsPageViewPage(){
		$res = $this->getResultsCommon(
			'ga:pageviews,ga:visits,ga:uniquePageviews,ga:avgTimeOnPage,ga:bounces,ga:exits,ga:bounceRate,ga:timeOnPage,ga:entrances',
			"ga:pagePath,ga:pageTitle",
			"-ga:pageviews"
		);

		//獲得ポイント=ページview数*ページview数/訪問数 * (100 - 直帰率)
		foreach($res["rows"] as $key => $v){
			if($v[3]>0){
				$res["rows"][$key]["point"] = $v[2] * ($v[2] / $v[3])*(100 - $v[6]/$v[3]);
			}else{
				$res["rows"][$key]["point"] = 0;
			}
			$point[] = $res["rows"][$key]["point"];
		}
		// 標準偏差
		$standard_deviation = $this->statsStandardDeviation($point);
		// 平均
		$avg = array_sum($point) / count($point);

		foreach($res["rows"] as $key => $v){
			//キーワード有効偏差値 = (獲得ポイント - 平均ポイント)*10/ポイント標準偏差 + 10
			$res["rows"][$key]["deviation"] = ($v["point"] - $avg)*10/$standard_deviation + 10;
		}

		return $res;
	}

	/*
	 * 累計
	 */
	function getResultsTotal(){
		$res = $this->getResultsCommon(
			'ga:pageviews,ga:visits,ga:organicSearches,ga:newVisits,ga:percentNewVisits,ga:exitRate,ga:bounceRate,ga:goalCompletionsAll,ga:users',
			"",
			""
			,1	
		);

		// organicsearches が期待通り（ga管理画面のorganic Searchの値）にならないので、再計算
		$search = $this->getResultSearch();
		if(isset($search["rows"][0][0])){
			$res["rows"][0][2] = $search["rows"][0][0];
		}

		return $res;
	}

	/*
	 * 累計（ユーザ種類ごと）
	 */
	function getResultsTotalUsertype(){
		$res = $this->getResultsCommon(
			'ga:pageviews,ga:visits,ga:organicSearches,ga:newVisits,ga:percentNewVisits,ga:exitRate,ga:bounceRate',
			"ga:userType",
			"",
			5	
		);
		return $res;
	}

	/*
	 * 累計（端末ごと）
	 */
	function getResultsTotalDevice(){
		$res = $this->getResultsCommon(
			'ga:pageviews,ga:visits,ga:organicSearches,ga:newVisits,ga:percentNewVisits,ga:exitRate,ga:bounceRate',
			"ga:deviceCategory",
			"",
			5	
		);
		return $res;
	}

	/*
	 * 累計（モバイル）
	 */
	function getResultsTotalMobile(){
		$res = $this->getResultsCommon(
			'ga:pageviews,ga:visits,ga:organicSearches,ga:newVisits,ga:percentNewVisits,ga:exitRate,ga:bounceRate',
			"",
			"",
			1,	
			"ga:deviceCategory==mobile,ga:deviceCategory==tablet"
		);

		return $res;
	}

	/*
	 * 累計（直帰）
	 */
	function getResultsTotalBounce(){
		$res = $this->getResultsCommon(
			'ga:entrances,ga:entranceRate,ga:bounces,ga:bounceRate',
			"",
			""
			,1	
		);
		return $res;
	}

	/*
	 * 検索セッション数
	 */
	function getResultSearch(){
		return $this->getResultsCommon(
			'ga:visits',
			"",
			"",
			1
			,'ga:medium==organic'
		);
	}

	/*
	 * 日別セッション数
	 */
	function getResultSearchDays(){
		return $this->getResultsCommon(
			'ga:visits',
			"ga:day",
			"ga:day",
			31,
			'ga:medium==organic'
		);
	}

	/*
	 * 特定ページの訪問数
	 */
	function getResultsAPageTotal($path){
		$res = $this->getResultsCommon(
			'ga:pageviews,ga:visits,ga:organicSearches,ga:newVisits,ga:percentNewVisits,ga:exitRate,ga:bounceRate',
			"",
			"",
			1,	
			"ga:pagePath==".$path
		);

		return $res;
	}

	/*
	 * コンバージョン
	 */
	function getGoalCompletionDays($conversion_id){
		return $this->getResultsCommon(
			'ga:goal'.$conversion_id.'Completions',
			"ga:day",
			"ga:day",
			31
		);
	}

	/*
	 * 年齢層別
	 */
	function getAgeBracket(){
		return $this->getResultsCommon(
			'ga:visits,ga:newVisits,ga:bounces,ga:pageviews,ga:avgSessionDuration,ga:goalConversionRateAll,ga:goalCompletionsAll',
			"ga:userAgeBracket",
			"ga:userAgeBracket",
			6
		);
	}

	/*
	 * トップページ詳細
	 */
	function getTopPage(){
		$res=array();
		$res["keyword"] = $this->getResultsCommon(
			'ga:pageviews,ga:visits,ga:bounces',
			"ga:Keyword",
			"-ga:pageviews",
			"10"
			,'ga:pagePath==/'
		);

		$res["previous"] = $this->getResultsCommon(
			'ga:pageviews',
			"ga:previousPagePath",
			"-ga:pageviews",
			"",
			"ga:previousPagePath!=(entrance);ga:previousPagePath!=/;ga:pagePath==/"
		);

		$res["second"] = $this->getResultsCommon(
			'ga:pageviews',
			"ga:secondPagePath",
			"-ga:pageviews",
			"",
			"ga:secondPagePath!=(not set);ga:secondPagePath!=/;ga:pagePath==/"
		);

		$res["start"] = $this->getResultsCommon(
			'ga:entrances,ga:entranceRate',
			"ga:pagePath",
			"-ga:entrances",
			""	
			,"ga:pagePath==/"
		);

		$res["end"] = $this->getResultsCommon(
			'ga:exits,ga:exitRate',
			"ga:pagePath",
			"-ga:exits",
			""	
			,"ga:pagePath==/"
		);

		return $res;
	}

	/*
	 * 標準偏差
	 */
	function statsStandardDeviation($ary) {
		// 平均取得
		$avg = array_sum($ary)/count($ary);

		// 各値の平均値との差の二乗【(値-平均値)^2】を算出
		$diff_ary = array();
		foreach ($ary as $val) {
			$diff = $val-$avg;
			$diff_ary[] = pow($diff,2);
		}

		// 上記差の二乗の合計を算出
		$diff_total = array_sum($diff_ary);
		// 平均を算出
		$diff_avg   = $diff_total/count($diff_ary);

		// 平方根を取る(標準偏差)
		$stdev = sqrt($diff_avg);

		// 標準偏差を返す
		return $stdev;
	}

}
