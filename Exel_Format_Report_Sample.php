<?php

set_include_path(get_include_path(). PATH_SEPARATOR .$_SERVER["DOCUMENT_ROOT"] . 'ga-api/PHPExcel/Classes/');

include_once 'PHPExcel.php';  
include_once 'PHPExcel/IOFactory.php';

class Exel_Format_Report{
	var $exel;
	var $sheet;
	var $file_name;
	var $save_dir;
	var $exel_type,$out_exel_type;

	function __construct(){

		$file="gar_evergreen_analytics_report.xls";
		$exel_type ="Excel5";

		$this->exel_type = $exel_type;
		$this->out_exel_type = $exel_type;
		$objReader = PHPExcel_IOFactory::createReader($this->exel_type);
		$objReader->setIncludeCharts(TRUE);
		$this->exel = $objReader->load($file);

		$this->save_dir = dirname(__FILE__)."/exel/";
	}

	// report
	/*
	 * 全体結果
	 */
	function resultTotal($gar){

		$init_month = $gar->month;
		$init_year = $gar->year;

		$ref = $gar->getResultsTotal();

		$this->setCellValue("C6", $gar->startDate);
		$this->setCellValue("E6", $gar->endDate);
		$this->setCellValue("H6", $gar->dayCount);

		$total= $ref["rows"][0];
		$this->setCellValue("E10", $total[0]); // 総ページビュー
		$this->setCellValue("E11", $total[1]); // 訪問者数
		$this->setCellValue("E12", round($total[0]/$total[1],1)); // 平均ページ閲覧数
		$this->setCellValue("E13", $total[2]); // 検索からの訪問数
		$this->setCellValue("E14", $this->percent($total[2]/$total[1])); // 検索からの訪問数の割合
		$this->setCellValue("E15", $total[3]); // 新規ユーザ数
		$this->setCellValue("E16", $this->percent($total[4]/100)); // 新規ユーザ数の割合

		// 先月までの月ごとデータ
		for($i=0;$i<22;$i++){

			$ref = $gar->getResultsTotal();

			$v = $ref["rows"][0];
			// 年月
			$this->setCellValue("B".($i+19), $gar->year."/".$gar->month);
			// ページビュー
			$this->setCellValue("C".($i+19), $v[0]);
			if((isset($next_month_0) && $v[0]>0)){
				// エクセル側のフォーマットに合わせる
				// 先月比
				$this->setCellValue("D".($i+18), $next_month_0/$v[0]);
			}
			// 訪問数
			$this->setCellValue("E".($i+19), $v[1]);
			if(isset($next_month_1) && $v[1]>0){
				// 先月比
				$this->setCellValue("F".($i+18), $next_month_1/$v[1]);
			}
			// 平均ページビュー
			$this->setCellValue("G".($i+19), ($v[1]>0) ? $v[0]/$v[1] : "-");
			if(isset($next_month_1)){
				// 先月比
				$this->setCellValue("H".($i+18), ($next_month_0/$next_month_1)/($v[0]/$v[1]));
			}
			// 検索訪問数
			$this->setCellValue("I".($i+19), $v[2]);
			// 検索訪問率
			$this->setCellValue("J".($i+19), ($v[1] > 0) ? $v[2]/$v[1] : "");
			if(isset($next_month_2) && $v[2]>0){
				// 先月比
				$this->setCellValue("K".($i+18), $next_month_2/$v[2]);
			}
			// 新規ユーザ数
			$this->setCellValue("L".($i+19), $v[3]);
			// 新規ユーザ率
			$this->setCellValue("M".($i+19), $v[4]/100);
			// 離脱率
			$this->setCellValue("N".($i+19), $v[5]/100);
			// 直帰率
			$this->setCellValue("O".($i+19), $v[6]/100);

			$next_month_0 = $v[0];
			$next_month_1 = $v[1];
			$next_month_2 = $v[2];

			if($i==1){
				$prev_month_pv = $v[0];
				$prev_month_visit = $v[1];
				$prev_month_per_pv = $v[1]/$v[0];
				$prev_month_search = $v[2];
				$prev_month_new_visit = $v[3];
			}
			if($i==2){
				$pprev_month_pv = $v[0];
				$pprev_month_visit = $v[1];
				$pprev_month_per_pv = $v[1]/$v[0];
				$pprev_month_search = $v[2];
				$pprev_month_new_visit = $v[3];
			}

			$prev_analyze = mktime(0,0,0,$gar->month-1,1,$gar->year);
			$prev_analyzd_year = date("Y",$prev_analyze);
			$prev_analyzd_month = date("n",$prev_analyze);
			$gar->setTerm($prev_analyzd_year,$prev_analyzd_month);
			
		}


		$this->setCellValue("G10", $this->percent($total[0]/$prev_month_pv)); // 総ページビュー
		$this->setCellValue("G11", $this->percent($total[1]/$prev_month_visit)); // 訪問者数
		$this->setCellValue("G12", $this->percent($prev_month_per_pv/($total[1]/$total[0]))); // 平均ページ閲覧数
		if($prev_month_search>0){
		$this->setCellValue("G13", $this->percent($total[2]/$prev_month_search)); // 検索からの訪問数
		}
		$this->setCellValue("G15", $this->percent($total[3]/$prev_month_new_visit)); // 新規ユーザ数
		$this->setCellValue("H10", $this->percent($total[0]/$pprev_month_pv)); // 総ページビュー
		$this->setCellValue("H11", $this->percent($total[1]/$pprev_month_visit)); // 訪問者数
		$this->setCellValue("H12", $this->percent($pprev_month_per_pv/($total[1]/$total[0]))); // 平均ページ閲覧数
		if($pprev_month_search>0){
		$this->setCellValue("H13", $this->percent($total[2]/$pprev_month_search)); // 検索からの訪問数
		}
		$this->setCellValue("H15", $this->percent($total[3]/$pprev_month_new_visit)); // 新規ユーザ数

		$gar->setTerm($init_year,$init_month);
	} 

	/*
	 * 日別レポート
	 */
	function resultDays($gar){
		$visiters = $gar->getResultsDays();

		$youbi = array("日","月","火","水","木","金","土");
		$youbi_sum = array();
		$youbi_cnt = array();

		foreach($visiters["rows"] as $key=>$v){
			$this->setCellValue("B".($key+7), $gar->year."/".$gar->month."/".$v[0]);
			$y=date("w",strtotime($gar->year."/".$gar->month."/".$v[0]));
			$this->setCellValue("C".($key+7), $youbi[$y]);
			$this->setCellValue("D".($key+7), $v[1]);
			$this->setCellValue("E".($key+7), $v[2]);
			$this->setCellValue("F".($key+7), round($v[1]/$v[2],1));

			if(!isset($youbi_sum[$y]["pv"])){
				$youbi_sum[$y]["pv"]=0;
			}
			if(!isset($youbi_sum[$y]["visit"])){
				$youbi_sum[$y]["visit"]=0;
			}
			$youbi_sum[$y]["pv"] += $v[1];
			$youbi_sum[$y]["visit"] += $v[2];

			if(!isset($youbi_cnt[$y])){
				$youbi_cnt[$y]=0;
			}
			$youbi_cnt[$y]++;
		}
		$this->setCellValue("D38", $visiters["sum"][1]);
		$this->setCellValue("E38", $visiters["sum"][2]);
		$this->setCellValue("F38", round($visiters["sum"][1]/$visiters["sum"][2],1));

		// 曜日別(月4回換算)
		for($i=0;$i<7;$i++){
			// 五回でてくる曜日は0.8倍する
			$rate = $youbi_cnt[$i]==5 ? 0.8 : 1;
			$pv = round($youbi_sum[$i]["pv"] * $rate , 0);
			$visit = round($youbi_sum[$i]["visit"] * $rate , 0);

			// 日曜が最後
			if($i==0){
				$this->setCellValue("I".($i+13), $pv);
				$this->setCellValue("J".($i+13), $visit);
				$this->setCellValue("K".($i+13), round($pv/$visit,2));
			}else{
				$this->setCellValue("I".($i+6), $pv);
				$this->setCellValue("J".($i+6), $visit);
				$this->setCellValue("K".($i+6), round($pv/$visit,2));
			}
		}

		$this->setCellValue("I14", $visiters["sum"][1]);
		$this->setCellValue("J14", $visiters["sum"][2]);
		$this->setCellValue("K14", round($visiters["sum"][1]/$visiters["sum"][2],1));

		// 時間別レポート
		$time = $gar->getResultsTime();
		foreach($time["rows"] as $key=>$v){
			$this->setCellValue("M".($key+7), $v[0]);
			$this->setCellValue("N".($key+7), $v[1]);
			$this->setCellValue("O".($key+7), $v[2]);
			$this->setCellValue("P".($key+7), round($v[2]/$v[1],2));
		}
		$this->setCellValue("N31", $time["sum"][1]);
		$this->setCellValue("O31", $time["sum"][2]);
		$this->setCellValue("P31", round($time["sum"][2]/$time["sum"][1],2));
	}

	/*
	 * 参照元
	 */
	function resultRefferral($gar){

		$this->setCellValue("B5", $gar->year."/".$gar->month);

		$ref = $gar->getResultsReferral();
		for($i=0;$i<50;$i++){
			$start_i = $i+8;
			if(isset($ref["rows"][$i])){
				$v = $ref["rows"][$i];
				$this->setCellValue("C".$start_i, $v[0]." / ".$v[1]);
				$this->setCellValue("D".$start_i, $v[2]);
				$this->setCellValue("E".$start_i, substr(round($v[7]/$v[2],3)*100 ,0, 4). "%");
				$this->setCellValue("F".$start_i, substr(round($v[8]/$v[10],3)*100 ,0, 4). "%");
			}else{
				$this->setCellValue("C".$start_i, "");
				$this->setCellValue("D".$start_i, "");
				$this->setCellValue("E".$start_i, "");
				$this->setCellValue("F".$start_i, "");
			}
		}

		$this->setCellValue("D6",$ref["sum"][2]);
		$this->setCellValue("E6",substr(round($ref["sum"][7]/$ref["sum"][2],3)*100 ,0, 4). "%");
		$this->setCellValue("F6",substr(round($ref["sum"][8]/$ref["sum"][10],3)*100 ,0, 4). "%");

	}

	/*
	 * 先月の参照元
	 */
	function resultRefferralPreviousMonth($gar){
		// 先月にセット
		$gar->setTerm($gar->year,$gar->month-1);

		$this->setCellValue("H5", $gar->year."/".$gar->month);

		$ref = $gar->getResultsReferral();
		for($i=0;$i<50;$i++){
			$start_i = $i+8;
			if(isset($ref["rows"][$i])){
				$v = $ref["rows"][$i];
				$this->setCellValue("I".$start_i, $v[0]." ".$v[1]);
				$this->setCellValue("J".$start_i, $v[2]);
				$this->setCellValue("K".$start_i, substr(round($v[7]/$v[2],3)*100 ,0, 4). "%");
				$this->setCellValue("L".$start_i, substr(round($v[8]/$v[10],3)*100 ,0, 4). "%");
			}else{
				$this->setCellValue("I".$start_i, "");
				$this->setCellValue("J".$start_i, "");
				$this->setCellValue("K".$start_i, "");
				$this->setCellValue("L".$start_i, "");
			}
		}

		$this->setCellValue("J6",$ref["sum"][2]);
		$this->setCellValue("K6",substr(round($ref["sum"][7]/$ref["sum"][2],3)*100 ,0, 4). "%");
		$this->setCellValue("L6",substr(round($ref["sum"][8]/$ref["sum"][10],3)*100 ,0, 4). "%");

		// 元に戻す
		$gar->setTerm($gar->year,$gar->month+1);
	}

	/*
	 * 検索キーワード
	 */
	function resultKeyword($gar){

		$this->setCellValue("B5", $gar->year."/".$gar->month);

		$ref = $gar->getResultsKeyword();

		for($i=0;$i<50;$i++){
			$start_i = $i+8;
			if(isset($ref["rows"][$i])){
				$v = $ref["rows"][$i];
				$this->setCellValue("C".$start_i, $v[0]);
				$this->setCellValue("D".$start_i, $v[2]);
				$this->setCellValue("E".$start_i, round($v[1]/$v[2],1));
//				$this->setCellValue("F".$start_i, substr(round($v[4]/$v[3],3)*100 ,0, 4). "%");
				$this->setCellValue("F".$start_i, substr(($v[8]/$v[7])*100 ,0, 4). "%");
			}else{
				$this->setCellValue("C".$start_i, "");
				$this->setCellValue("D".$start_i, "");
				$this->setCellValue("E".$start_i, "");
				$this->setCellValue("F".$start_i, "");
			}
		}

		$this->setCellValue("D6",$ref["sum"][2]);
		$this->setCellValue("E6",round($ref["sum"][1]/$ref["sum"][2],1));
//		$this->setCellValue("F6",substr(round($ref["sum"][4]/$ref["sum"][3],3)*100 ,0, 4). "%");
		$this->setCellValue("F6",substr(round($ref["sum"][8]/$ref["sum"][7],3)*100 ,0, 4). "%");
	}

	/*
	 * 先月の検索キーワード
	 */
	function resultKeywordPreviousMonth($gar){
		// 先月にセット
		$gar->setTerm($gar->year,$gar->month-1);

		$this->setCellValue("H5", $gar->year."/".$gar->month);

		$ref = $gar->getResultsKeyword();

		for($i=0;$i<50;$i++){
			$start_i = $i+8;
			if(isset($ref["rows"][$i])){
				$v = $ref["rows"][$i];
				$this->setCellValue("I".$start_i, $v[0]);
				$this->setCellValue("J".$start_i, $v[2]);
				$this->setCellValue("K".$start_i, round($v[1]/$v[2],1));
//				$this->setCellValue("L".$start_i, substr(round($v[4]/$v[3],3)*100 ,0, 4). "%");
				$this->setCellValue("L".$start_i, substr(($v[8]/$v[7]) * 100 ,0, 4). "%");
			}else{
				$this->setCellValue("I".$start_i, "");
				$this->setCellValue("J".$start_i, "");
				$this->setCellValue("K".$start_i, "");
				$this->setCellValue("L".$start_i, "");
			}
		}

		$this->setCellValue("J6",$ref["sum"][2]);
		$this->setCellValue("K6",round($ref["sum"][1]/$ref["sum"][2],1));
//		$this->setCellValue("L6",substr(round($ref["sum"][4]/$ref["sum"][3],3)*100 ,0, 4). "%");
		$this->setCellValue("L6",substr(($ref["sum"][8]/$ref["sum"][7])*100 ,0, 4). "%");

		// 元に戻す
		$gar->setTerm($gar->year,$gar->month+1);

	}

	/*
	 * 直帰ページ
	 */
	function resultBouncePage($gar){

		$this->setCellValue("B5", $gar->year."/".$gar->month);

		$ref = $gar->getResultsBouncePage();
		for($i=0;$i<25;$i++){
			$start_i = $i+8;
			if(isset($ref["rows"][$i])){
				$v = $ref["rows"][$i];
				$this->setCellValue("C".$start_i, $v[0] ."\n" .$v[1]);
				$this->sheet->getStyle('C'.$start_i)->getAlignment()->setWrapText(true);
				$this->setCellValue("D".$start_i, $v[2]);
				$this->setCellValue("E".$start_i, $v[3]);
				$this->setCellValue("F".$start_i, $this->percent($v[3]/$v[2],true));
			}else{
				$this->setCellValue("C".$start_i, "");
				$this->setCellValue("D".$start_i, "");
				$this->setCellValue("E".$start_i, "");
				$this->setCellValue("F".$start_i, "");
			}
		}

		$this->setCellValue("D6",$ref["sum"]["2"]);
		$this->setCellValue("E6",$ref["sum"]["3"]);
		$this->setCellValue("F6",$this->percent($ref["sum"]["3"]/$ref["sum"]["2"],true));

	}

	/*
	 * 先月の直帰ページ
	 */
	function resultBouncePagePreviousMonth($gar){

		// 先月
		$gar->setTerm($gar->year,$gar->month-1);

		$this->setCellValue("H5", $gar->year."/".$gar->month);

		$ref = $gar->getResultsBouncePage();
		for($i=0;$i<25;$i++){
			$start_i = $i+8;
			if(isset($ref["rows"][$i])){
				$v = $ref["rows"][$i];
				$this->setCellValue("I".$start_i, $v[0] ."\n" .$v[1]);
				$this->sheet->getStyle('I'.$start_i)->getAlignment()->setWrapText(true);
				$this->setCellValue("J".$start_i, $v[2]);
				$this->setCellValue("K".$start_i, $v[3]);
				$this->setCellValue("L".$start_i, $this->percent($v[3]/$v[2],true));
			}else{
				$this->setCellValue("I".$start_i, "");
				$this->setCellValue("J".$start_i, "");
				$this->setCellValue("K".$start_i, "");
				$this->setCellValue("L".$start_i, "");
			}
		}

		$this->setCellValue("J6",$ref["sum"]["2"]);
		$this->setCellValue("K6",$ref["sum"]["3"]);
		$this->setCellValue("L6",$this->percent($ref["sum"]["3"]/$ref["sum"]["2"],true));

		// 元に戻す
		$gar->setTerm($gar->year,$gar->month+1);
	}

	/*
	 * 閲覧開始ページ
	 */
	function resultEntrancePage($gar){

		$this->setCellValue("B5", $gar->year."/".$gar->month);

		$ref = $gar->getResultsEntrancePage();

		for($i=0;$i<25;$i++){
			$start_i = $i+8;
			if(isset($ref["rows"][$i])){
				$v = $ref["rows"][$i];
				$this->setCellValue("C".$start_i, $v[0] ."\n" .$v[1]);
				$this->sheet->getStyle('C'.$start_i)->getAlignment()->setWrapText(true);
				$this->setCellValue("D".$start_i, $v[4]);
				if($v[4]>0){
					$this->setCellValue("E".$start_i, $this->percent($v[5]/$v[4],true));
				}else{
					$this->setCellValue("E".$start_i, "--%");
				}
			}else{
				$this->setCellValue("C".$start_i, "");
				$this->setCellValue("D".$start_i, "");
				$this->setCellValue("E".$start_i, "");
			}
		}

		$this->setCellValue("D6",$ref["sum"][4]);
		$this->setCellValue("E6",$this->percent($ref["sum"][5]/$ref["sum"][4],true));

	}

	/*
	 * 先月の閲覧開始ページ
	 */
	function resultEntrancePagePreviousMonth($gar){

		// 先月にセット
		$gar->setTerm($gar->year,$gar->month-1);

		$this->setCellValue("G5", $gar->year."/".$gar->month);

		$ref = $gar->getResultsEntrancePage();
		for($i=0;$i<25;$i++){
			$start_i = $i+8;
			if(isset($ref["rows"][$i])){
				$v = $ref["rows"][$i];
				$this->setCellValue("H".$start_i, $v[0] ."\n" .$v[1]);
				$this->sheet->getStyle('H'.$start_i)->getAlignment()->setWrapText(true);
				$this->setCellValue("I".$start_i, $v[4]);
				if($v[3]>0){ // 訪問数が0のとき、開始数も離脱数も0になる対応
					$this->setCellValue("J".$start_i, $this->percent($v[5]/$v[4],true));
				}else{
					$this->setCellValue("J".$start_i, "--%");
				}
			}else{
				$this->setCellValue("H".$start_i, "");
				$this->setCellValue("I".$start_i, "");
				$this->setCellValue("J".$start_i, "");
			}
		}

		$this->setCellValue("I6",$ref["sum"][4]);
		$this->setCellValue("J6",substr(round($ref["sum"][5]/$ref["sum"][4],3)*100 ,0, 4). "%");

		// 元に戻す
		$gar->setTerm($gar->year,$gar->month+1);

	}

	/*
	 * 離脱ページ
	 */ 
	function resultExitPage($gar){

		$this->setCellValue("B5", $gar->year."/".$gar->month);

		$ref = $gar->getResultsExitPage();
		for($i=0;$i<25;$i++){
			$start_i = $i+8;
			if(isset($ref["rows"][$i])){
				$v = $ref["rows"][$i];
				$this->setCellValue("C".$start_i, $v[0] ."\n" .$v[1]);
				$this->sheet->getStyle('C'.$start_i)->getAlignment()->setWrapText(true);
				$this->setCellValue("D".$start_i, $v[3]);
				$this->setCellValue("E".$start_i, $this->percent($v[4]/100,true));
			}else{
				$this->setCellValue("C".$start_i, "");
				$this->setCellValue("D".$start_i, "");
				$this->setCellValue("E".$start_i, "");
			}
		}

		$this->setCellValue("D6",$ref["sum"][3]);
		$this->setCellValue("E6",$this->percent($ref["sum"][3]/$ref["sum"][2],true));
	}

	/*
	 * 先月の離脱ページ
	 */
	function resultExitPagePreviousMonth($gar){
		// 先月
		$gar->setTerm($gar->year,$gar->month-1);

		$this->setCellValue("G5", $gar->year."/".$gar->month);

		$ref = $gar->getResultsExitPage();
		for($i=0;$i<25;$i++){
			$start_i = $i+8;
			if(isset($ref["rows"][$i])){
				$v = $ref["rows"][$i];
				$this->setCellValue("H".$start_i, $v[0] ."\n" .$v[1]);
				$this->sheet->getStyle('H'.$start_i)->getAlignment()->setWrapText(true);
				$this->setCellValue("I".$start_i, $v[3]);
				$this->setCellValue("J".$start_i, $this->percent($v[4]/100,true));
			}else{
				$this->setCellValue("H".$start_i, "");
				$this->setCellValue("I".$start_i, "");
				$this->setCellValue("J".$start_i, "");
			}
		}

		$this->setCellValue("I6",$ref["sum"][3]);
		$this->setCellValue("J6",$this->percent($ref["sum"][3]/$ref["sum"][2],true));

		// 元に戻す
		$gar->setTerm($gar->year,$gar->month+1);

	}

	/*
	 * 閲覧ページ
	 */
	function resultViewPage($gar){

		$this->setCellValue("B5", $gar->year."/".$gar->month);

		$ref = $gar->getResultsPageViewPage();
		for($i=0;$i<25;$i++){
			$start_i = $i+8;
			if(isset($ref["rows"][$i])){
				$v = $ref["rows"][$i];
				$this->setCellValue("C".$start_i, $v[0] ."\n" .$v[1]);
				$this->sheet->getStyle('C'.$start_i)->getAlignment()->setWrapText(true);

				$this->setCellValue("D".$start_i, $v[2]);
//				$this->setCellValue("E".$start_i, $v[3]);
				$this->setCellValue("E".$start_i, $v[4]);
			}else{
				$this->setCellValue("C".$start_i, "");
				$this->setCellValue("D".$start_i, "");
				$this->setCellValue("E".$start_i, "");
			}
		}

		$this->setCellValue("D6",$ref["sum"][2]);
//		$this->setCellValue("E6",$ref["sum"][3]);
		$this->setCellValue("E6",$ref["sum"][4]);

	}

	/*
	 * 先月の閲覧ページ
	 */
	function resultViewPagePrevioustMonth($gar){
		// 先月
		$gar->setTerm($gar->year,$gar->month-1);

		$this->setCellValue("G5", $gar->year."/".$gar->month);

		$ref = $gar->getResultsPageViewPage();
		for($i=0;$i<25;$i++){
			$start_i = $i+8;
			if(isset($ref["rows"][$i])){
				$v = $ref["rows"][$i];
				$this->setCellValue("H".$start_i, $v[0] ."\n" .$v[1]);
				$this->sheet->getStyle('H'.$start_i)->getAlignment()->setWrapText(true);

				$this->setCellValue("I".$start_i, $v[2]);
				$this->setCellValue("J".$start_i, $v[4]);
			}else{
				$this->setCellValue("H".$start_i, "");
				$this->setCellValue("I".$start_i, "");
				$this->setCellValue("J".$start_i, "");
			}
		}

		$this->setCellValue("I6",$ref["sum"][2]);
		$this->setCellValue("J6",$ref["sum"][4]);

		// 元に戻す
		$gar->setTerm($gar->year,$gar->month+1);

	}

	// ****
	// basic
	// ****
	
	function setSheet($index){
		// 全体
		$this->exel->setActiveSheetIndex($index);
		$this->sheet = $this->exel->getActiveSheet();
	}

	function setCellValue($cell,$val){
		$this->sheet->setCellValue($cell, $val);
	}

	function setFileName($file_name){
		$this->file_name = $file_name;
	}

	function setOutExelType($exel_type){
		$this->out_exel_type = $exel_type;
	}

	function save(){
		if($this->file_name==""){
			die('filename is empty');
		}

		// エクセル出力
		$writer = PHPExcel_IOFactory::createWriter($this->exel, $this->out_exel_type);
		$writer->setIncludeCharts(TRUE);
		$writer->save($this->save_dir . $this->file_name);
	}

	function out(){
		// Redirect output to a client’s web browser (Excel5)
		header('Content-Type: application/vnd.ms-excel');
		header('Content-Disposition: attachment;filename="'.$this->file_name.'"');
		header('Cache-Control: max-age=0');
		//// If you're serving to IE 9, then the following may be needed
		//header('Cache-Control: max-age=1');
		// If you're serving to IE over SSL, then the following may be needed
		header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
		header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
		header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
		//header ('Pragma: public'); // HTTP/1.0

		$writer = PHPExcel_IOFactory::createWriter($this->exel, $this->out_exel_type);
		$writer->setIncludeCharts(TRUE);
		$writer->save('php://output');
	}

	function percent($v,$str_percent=false){
		return sprintf("%.1f", $v * 100).($str_percent ? "%":"");
	}

	// ******
	// メイン処理
	// ******
	function make($gar){
		// 全体
		$this->setSheet(1);
		$this->resultTotal($gar);

		// 日別レポート
		$this->setSheet(2);
		$this->resultDays($gar);

		// 参照元
		$this->setSheet(3);
		$this->resultRefferral($gar);
		$this->resultRefferralPreviousMonth($gar);

		// 検索キーワード
		$this->setSheet(4);
		$this->resultKeyword($gar);
		$this->resultKeywordPreviousMonth($gar);

		// 閲覧開始ページ
		$this->setSheet(5);
		$this->resultEntrancePage($gar);
		$this->resultEntrancePagePreviousMonth($gar);

		// 離脱ページ
		$this->setSheet(6);
		$this->resultExitPage($gar);
		$this->resultExitPagePreviousMonth($gar);

		// 直帰ページ
		$this->setSheet(7);
		$this->resultBouncePage($gar);
		$this->resultBouncePagePreviousMonth($gar);

		// 閲覧ページ
		$this->setSheet(8);
		$this->resultViewPage($gar);
		$this->resultViewPagePrevioustMonth($gar);
		
		// レポートタイトル
		$this->setSheet(0);
		$this->setCellValue("N14", $gar->getProfileName());
		$this->setCellValue("N16", $gar->startDate."-".$gar->endDate);

		$this->setFileName(str_replace("/","",$gar->getProfileName())."_".$gar->year."_".$gar->month.".xls");
	}
}
