<?php
set_include_path(dirname(__FILE__)."/google-api-php-client/src" . PATH_SEPARATOR . get_include_path());
require_once "Google/Client.php";
require_once "Google/Service/Analytics.php";
require_once dirname(__FILE__)."/config.php";

class Google_Analytics_Report_Core{

	/*
	 * 1リクエストのページングの件数の上限
	 */
	const API_LIMIT_PER_REQ=10000;

	/*
	 * Google_Client
	 */
	var $client;

	/*
	 * analytics
	 */
	var $analytics;

	/*
	 * accounts
	 */
	var $accounts;

	/*
	 * profileId
	 */
	var $profileId;

	/*
	 * profileName
	 */
	var $profileName;

	/*
	 * year
	 */
	var $year;

	/*
	 * month
	 */
	var $month;

	/*
	 * day
	 */
	var $day;

	/*
	 * term
	 */
	var $term;

	/*
	 * Terms of start analyzing date
	 */
	var $startDate;

	/*
	 * Terms of end analyzing date
	 */
	var $endDate;

	/*
	 * Date of start analytics
	 */
	var $startTermDate;

	/*
	 * day count
	 */
	var $dayCount;

	function __construct(){
		session_start();
		$this->client = new Google_Client();

		$this->client->setClientId(GAR_API_CLIENT_ID);
		$this->client->setClientSecret(GAR_API_CLIENT_SECRET);
		$this->client->setDeveloperKey(GAR_API_DEVELOPER_KEY);
		$this->client->setRedirectUri(GAR_API_REDIRECT_URI);
		$this->client->setScopes(array('https://www.googleapis.com/auth/analytics'));
		
		if (isset($_SESSION['google_analytics_report_token']) && $_SESSION['google_analytics_report_token'] != "") {
			$this->client->setAccessToken($_SESSION['google_analytics_report_token']);
		}else{
			$authUrl = $this->client->createAuthUrl();

			print "<a class='login' href='$authUrl'>login google account</a>";
			exit;
		}
		
		if($this->client->isAccessTokenExpired()) {
			//    echo 'Access Token Expired'; // Debug


			$authUrl = $this->client->createAuthUrl();

			print "<a class='login' href='$authUrl'>login google account</a>";
			exit;
		}

		$this->analytics = new Google_Service_Analytics($this->client);
	}

	function setProfileId($profileId){
		$this->profileId = $profileId;
	}

	function getProfileId(){
		return $this->profileId;
	}

	function setProfileName($name){
		$this->profileName=$name;
	}

	function getProfileName(){
		return $this->profileName;
	}

	function getAccounts(){
		return $this->accounts;
	}
		 
	/*
	 * 解析期間指定
	 */
	function setTerm($year,$month,$day=1,$term="month"){
		$this->year = $year;
		$this->month = $month;
		$this->day = $day;
		$this->term = $term;

		if($term=="month"){
			$this->dayCount = date("t",mktime(0,0,0,$this->month,$this->day,$this->year));
		}elseif($term=="week"){
			$this->dayCount = 7;
		}

		$this->setDefaultTerm();
	}

	/*
	 * 解析開始期間
	 */
	function setStartTerm($year,$month,$day=1){
		$this->startTermDate = mktime(0,0,0, $month, $day, $year);
	}

	/*
	 * 前の期間をセット
	 */
	function setPrevTerm(){
			$t = strtotime($this->startDate);
		if($this->term=="month"){
			$this->startDate = date('Y/m/1',$t - 60*60*24*15);
			$this->endDate = date('Y/m/t',$t - 60*60*24*15);
		}elseif($this->term=="week"){
			$t = strtotime($this->startDate);
			$this->startDate = date('Y/m/d',$t - 60*60*24*7);
			$this->endDate = date('Y/m/d',$t - 60*60*24*1);
		}

	}

	/*
	 * 今の期間に戻す
	 */
	function setDefaultTerm(){

		// 週単位の場合
		if($this->term=="week"){
			$this->startDate = $this->year."/".$this->month."/".$this->day;
			$this->endDate = date('Y/m/d', mktime(0,0,0, $this->month, $this->day+6, $this->year));
		}elseif($this->term=="month"){
			$this->startDate = $this->year."/".$this->month."/1";
			$this->endDate = date('Y/m/t', mktime(0,0,0, $this->month, 1, $this->year));
		}else{
			die("execpt setTerm()");
		}
	}

	/*
	 * プロフィールID取得
	 */
	function setAccounts(){

		$accounts = $this->analytics->management_accounts->listManagementAccounts();

		$ext_cnt = 0;

		if (count($accounts->getItems()) > 0) {
			$items = $accounts->getItems();
			foreach($items as $key_a => $item){

				$accountId = $item->getId();
				$webproperties = $this->analytics->management_webproperties
					->listManagementWebproperties($accountId);

				if (count($webproperties->getItems()) > 0) {
					$items_w = $webproperties->getItems();
					foreach($items_w as $key_w => $item_w){
						$webpropertyId = $item_w->getId();

						$profiles = $this->analytics->management_profiles
							->listManagementProfiles($accountId, $webpropertyId);

						if (count($profiles->getItems()) > 0) {
							$items_p = $profiles->getItems();
							foreach($items_p as $key_p => $item_p){
								$this->accounts[] = $item_p;
							}

						} else {
							$ext_cnt++;
							// プロフィールが多すぎるとエラーがでるの、ある程度(10)で止める
							if($ext_cnt>10) return;
						}
					}

				} else {
					throw new Exception('No webproperties found for this user.');
				}
			}
		} else {
			throw new Exception('No accounts found for this user.');
		}
	}

	
	//***************
	// report format
	//***************

	function getResultsCommon($metrics,$dimensions,$sort,$max_results=0,$filters="",$segment=""){

		// プロフィールIDチェック
		if(!$this->profileId){
			die("profile ID not set");
		}

		// 各パラメータ設定
		$optParams = array(
			'dimensions' => $dimensions,
		);

		if($sort!=""){
			$optParams['sort'] = $sort;
		}
		if($filters!=""){
			$optParams['filters'] = $filters;
		}
		if($segment!=""){
			$optParams['segment'] = $segment;
		}

		// 最大件数が設定されてて取得件数が１ページ件数内の場合、負荷軽減のためmax_resultsに抑える
		if($max_results>0 && $max_results < self::API_LIMIT_PER_REQ){
			$optParams['max-results'] = $max_results;
		}else{
			$optParams['max-results'] = self::API_LIMIT_PER_REQ;
		}

		// apiよりデータ取得
		$rows = array();
		$cnt = 0;
		$max_result_flg=false;
		$sd = date("Y-m-d",strtotime($this->startDate));
		$ed = date("Y-m-d",strtotime($this->endDate));

		for($i=1;;$i+=self::API_LIMIT_PER_REQ){
			$optParams['start-index'] = $i;

			// 最大件数が設定されてて、最大件数以上を取得しようとした場合
			if($max_results>0 && $i > $max_results) break;

			try{
				$res = $this->analytics->data_ga->get(
					'ga:' . $this->profileId,
					$sd,
					$ed,
					$metrics,
					$optParams
				);

				// 解析日以外だと取れないのでその対応
				if(!is_array($res->getRows())) {
					break;
				}

				foreach($res->getRows() as $v){
					// データセット 
					$rows[$cnt] = $v;
					$cnt++;

					// 最大件数が設定されてて、最大件数以上を取得しようとした場合
					if($max_results > 0 && $cnt >= $max_results){
						$max_result_flg=true;
						break;
					}
				}
				// 最大件数に到達した場合
				if($max_result_flg) break;

				// ページング件数取得できなかった==最後のページになった
				if(count($res->getRows()) < self::API_LIMIT_PER_REQ) break;

			}catch(Exception $e){
				// データ取得のエラー処理
				var_dump($e->getCode());
				var_dump($e->getMessage());
				exit;
			}
		}

		$result=array();

		// 結果
		$result["rows"] = $rows;
		// 取得件数
		$result["cnt"] = $cnt;
		// 各値の合計値
		$result["sum"] = array();
		if(is_array($rows) && count($rows)>0){
			foreach($rows as $key=>$val){
				foreach($val as $k=>$v){
					if(is_numeric($v)){
						if(!isset($result["sum"][$k])) $result["sum"][$k]=0;
						$result["sum"][$k] += $v;
					}
				}
			}
		}

		// プロフィール名
		if(isset($res)){
			if(!$this->profileName){
				$this->setProfileName($res->getProfileInfo()->getProfileName());
			}
		}

		return $result;

	}

}
