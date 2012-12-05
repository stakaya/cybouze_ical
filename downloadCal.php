<?php
/*********************************************************************************/
/**
 *  概要：サイボウズからスケジュールをダウンロードし、ical形式に変換します。
 *  使い方:下記のコマンドを実行します。
 *         コマンドライン引数にサイボウズのIDとパスワードを指定します。  
 * 
 *  php -f downloadCal.php サイボウズのid パスワード 
 */

// サイボウズカレンダーのアドレス 
define('CAL_URL', 'cbz.hogehoge.co.jp');

// 未来何ヶ月分取得するか↓数字部分で調整
define('CAL_TERM', '+2 month');

/*********************************************************************************/ 

if ($argc != 3) {
	die('パラメータエラー');
}

require_once 'iCalcreator.class.php';

mb_internal_encoding('UTF-8');     
date_default_timezone_set('Asia/Tokyo');

function httpPost($url, $file, $cookie = null, $params = null, $port = 80, $timeout = 10) {

	$fp = @fsockopen(basename($url), $port, $errno, $errstr, $timeout);
	if (!$fp) {
		return false;
	} else {
		$req  = "POST $file HTTP/1.0\r\n";
		$req .= "Host: " . basename($url) . "\r\n";
		$req .= "Content-Type: application/x-www-form-urlencoded\r\n";
		$req .= "Content-Length: " . strlen($params) . "\r\n";
		$req .= "User-Agent: MSIE7.0\r\n";
		if ($cookie != null) {
			$req .= 'Cookie:';
			$AllCookies = ''; 
			for ($i = 0; $i < count($cookie); $i++) {
				$temp = substr($cookie[$i], 12);
				$temp = substr($temp, 0, strpos($temp, ' '));
				$AllCookies .= ' '. $temp;
			}
			$req .= $AllCookies . "\r\n";
		}

		$req .= "Connection: Close\r\n\r\n";
		if ($params != null) {
			$req .= $params . "\r\n";
		}

		fwrite($fp, $req);
		$temp = '';
		while (!feof($fp)) {
			$temp .= fgets($fp, 1024);
		}
		fclose($fp);

		// HTTPヘッダを取り除く
		return substr($temp, strpos($temp, "\r\n\r\n") +4);
	}
}

function httpGet($url, $file, $cookie = null, $params = null, $port = 80, $timeout = 10) {

	$fp = @fsockopen(basename($url), $port, $errno, $errstr, $timeout);
	if (!$fp) {
		return false;
	} else {

		$req  = "GET $file HTTP/1.0\r\n";
		$req .= "Host: " . basename($url) . "\r\n";

		if ($cookie != null) {
			$req .= 'Cookie:';
			$AllCookies = ''; 
			for ($i = 0; $i < count($cookie); $i++) {
				$temp = substr($cookie[$i], 12);
				$temp = substr($temp, 0, strpos($temp, ' '));
				$AllCookies .= ' '. $temp;
			}
			$req .= $AllCookies . "\r\n";
		}

		$req .= "Connection: Close\r\n\r\n";
		if ($params != null) {
			$req .= $params . "\r\n";
		}

		fwrite($fp, $req);
		$temp = '';
		while (!feof($fp)) {
			$temp .= fgets($fp, 1024);
		}
		fclose($fp);

		return $temp;
	}
}

$response = httpGet(CAL_URL, '/cgi-bin/cbag/ag.cgi?_ID='.$argv[1].'&Password='.$argv[2].'&_System=login&_Login=1&GuideNavi=1');
$ret = preg_match_all('/Set-Cookie: .*\n/', $response, $cookie);
if ($ret == 0) {
	die('認証エラー');
}

$params = array(
	'page' => 'PersonalScheduleExport',
	'SetDate.Year' => date('Y'),
	'SetDate.Month' =>intval(date('m')),
	'SetDate.Day' =>intval(date('d')),
	'EndDate.Year' => date('Y', strtotime(CAL_TERM)),
	'EndDate.Month' => intval(date('m', strtotime(CAL_TERM))),
	'EndDate.Day' => intval(date('d', strtotime(CAL_TERM))),
	'ItemName' => '0',
	'Export' => urlencode('書き出す')
);

$fileName = $argv[1] . '.schedule.csv'; 
$temp = httpPost(CAL_URL, '/cgi-bin/cbag/ag.cgi/schedule.csv?', $cookie[0], http_build_query($params));
$fp = fopen($fileName, 'wb');
@fwrite($fp, $temp, strlen($temp));
fclose($fp);

$file = fopen($fileName, 'r'); 

if(!feof($file)){ 
	$temp = fgetcsv($file);
}	

$v = new vcalendar(array('unique_id' => 'cybouze.calender'));

while(!feof($file)){ 
	$temp = fgetcsv($file);
	mb_convert_variables('UTF-8', 'SJIS-win', $temp); 

	if(count($temp) != 9) {
		continue;
	}

	$e =& $v->newComponent('vevent');

	$date = explode('/', $temp[0]);
	$time = explode(':', $temp[1]);
	if(count($time) == 3) {
		$e->setProperty('dtstart',
			$date[0], $date[1], $date[2],
			$time[0], $time[1], $time[2]);
	} else {
		$e->setProperty('dtstart',
			$date[0], $date[1], $date[2]);
	}

	$date = explode('/', $temp[2]);
	$time = explode(':', $temp[3]);
	if(count($time) == 3) {
		$e->setProperty('dtend',
			$date[0], $date[1], $date[2],
			$time[0], $time[1], $time[2]);
	} else {
		$e->setProperty('dtend', $e->getProperty('dtstart'));
	}

	if (strlen($temp[4]) > 0) {
		$e->setProperty('summary', $temp[4]);
	} else {
		$e->setProperty('summary', $temp[5]);
	}	
	$e->setProperty('description', $temp[5]);
	$e->setProperty('location', $temp[8]);             

	$e->setProperty('class', 'PRIVATE');
	$e->setProperty('categories', 'BUSINESS'); 
} 

fclose($file);

$temp = $v->createCalendar();
print mb_convert_encoding($temp, 'sjis-win');

$fp = fopen($argv[1] .'.cybouze.calender.ics', 'w');
@fwrite($fp, $temp, strlen($temp));
fclose($fp);
?>
