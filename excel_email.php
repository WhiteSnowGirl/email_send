<html>
 <head> 
 <title></title>  
 <meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
 <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1"> 
 <link rel="stylesheet" type="text/css" href="static/style.css" />
 <script type="text/javascript" src="static/jquery-1.8.2.min.js"></script>
 </head>
 <body> 
 <div  style="text-align: center;margin-top: 60px;">
 	<div> <h1>Email-send</h1></div>

 <div class="main">
<form id="form1" name="form1" method="post" action="">  

  	
  	<div class="div1">
	    <div class="div2">上传文件</div>
	    <input name="file" type="file" class="inputstyle" id="file"/> 
	    
	    <div class="filename"></div>
	</div>
  	<input class="tijiao" type="submit" name="Submit" value="提交" />  
 
</form> 
</div>
</div>
<script>
$(document).ready(function(){
	$('.sucess').fadeOut(3000);
	$('.unread').fadeOut(2000);
	$('#file').change(function(){
		var name =$('#file').val();
		var arr=name.split('\\');//注split可以用字符或字符串分割 
		var name=arr[arr.length-1];
		$('.filename').html(name);
	});
	
	
	// var arr=name.split('\\');//注split可以用字符或字符串分割 
	// var my=arr[arr.length-1];//这就是要取得的图片名称 
	// alert(my);
});
</script>
</body>
</html>   
<?php
error_reporting(E_STRICT);
date_default_timezone_set('Asia/Shanghai');
// 请求 PHPmailer类 文件
require_once("class.phpmailer.php");
require_once 'reader.php';    
$data = new Spreadsheet_Excel_Reader();   
$data->setOutputEncoding('gbk');  
//发送Email函数
function smtp_mail ( $sendto_email, $subject, $body, $extra_hdrs, $user_name) {
$mail = new PHPMailer(); 
$mail->IsSMTP(); // send via SMTP 
$mail->Host = "smtp.163.com";   // SMTP servers 
$mail->SMTPAuth = true; // turn on SMTP authentication 
$mail->Username = "zz@163.com";  // SMTP username 注意：普通邮件认证不需要加 @域名
$mail->Password = "123456"; // SMTP password

$mail->From = "zz@163.com";  // 发件人邮箱
$mail->FromName = 人事行政总监; //   发件人 ,名称
$mail->CharSet = "GB2312";  // 这里指定字符集！
$mail->Encoding = "base64";

$mail->AddAddress($sendto_email,$user_name);// 收件人邮箱和姓名
$mail->AddReplyTo("","jingjing_test");

//$mail->WordWrap = 50; // set word wrap 
//$mail->AddAttachment("/var/tmp/file.tar.gz");// attachment  附件1
//$mail->AddAttachment("/tmp/image.jpg", "new.jpg"); //附件2
$mail->IsHTML(true);   // send as HTML 
$mail->Subject = $subject;  
$mail->Body = $body;
$mail->AltBody ="text/html"; 
$date= date('Y-m-d H:i:s');
if($mail->Send()) 
{ 
   info_write("ok.txt","$user_name 发送成功 $date");
} 
else {
   info_write("falied.txt","$user_name 失败,错误信息$mail->ErrorInfo  $date");
 }
}
// 发送Email函数结束

// 写入发送结果函数，错误日志记录
function info_write($filename,$info_log)
{
 $info.= $info_log;
 $info.="\r\n";
 $fp = fopen ($filename,a);
 fwrite($fp,$info);
 fclose($fp);
}

//定时跳转页面 函数其中 1000是时间,1秒, 您可以自定义
function redirect($url)
{
echo "<script>
function redirect()
{
window.location.replace('$url');
}
window.setTimeout('redirect();', 1000);
  </script>";
}

//读取文本 邮件地址  您也可以读 数据库
/*$filename = "email.txt";
$fp = fopen($filename,"r");
$contents = fread($fp,filesize($filename));
$list_email=explode("\r\n",$contents); 
$len=count($list_email);
fclose($fp);*/
// 参数说明(发送到, 邮件主题, 邮件内容, 附加信息, 用户名)
$i = $_GET['action']; 
if($_POST['Submit'])  
{ 
$data->read($_POST['file']);  
for ($i = 1; $i <= $data->sheets[0]['numRows']; $i++) {
$list_email=$data->sheets[0]['cells'][$i][1]; //邮箱地址
$company_name=$data->sheets[0]['cells'][$i][2];//公司名称
$department=$data->sheets[0]['cells'][$i][3];//一级部门
$job_num=$data->sheets[0]['cells'][$i][4];//工号
$name=$data->sheets[0]['cells'][$i][5];//员工姓名
$work_allday=$data->sheets[0]['cells'][$i][6];//应出勤
$work_realday=$data->sheets[0]['cells'][$i][7];//实出勤
$business_travel=$data->sheets[0]['cells'][$i][8];//出差
$private_leave =$data->sheets[0]['cells'][$i][9];//事假
$sick_leave=$data->sheets[0]['cells'][$i][10];//病假
$marital_leave=$data->sheets[0]['cells'][$i][11];//婚假
$maternity_leave=$data->sheets[0]['cells'][$i][12];//产假
$maternity_leave_company=$data->sheets[0]['cells'][$i][13];//陪产假
$jbereavement_leave=$data->sheets[0]['cells'][$i][14];//丧假
$annual_leave=$data->sheets[0]['cells'][$i][15];//年休假
$statutory_holidays=$data->sheets[0]['cells'][$i][16];//法定节假日
$statutory_holiday_overtime=$data->sheets[0]['cells'][$i][17];//法定节假日加班
$weekend_overtime=$data->sheets[0]['cells'][$i][18];//周末加班
$working_overtime=$data->sheets[0]['cells'][$i][19];//工作日加班
$be_on_duty=$data->sheets[0]['cells'][$i][20];//值班
$neglect_work=$data->sheets[0]['cells'][$i][21];//旷工
$wage=$data->sheets[0]['cells'][$i][22];//工资标准
$attendance_standards=$data->sheets[0]['cells'][$i][23];//出勤工资
$overtime_wages=$data->sheets[0]['cells'][$i][24];//加班工资
$meal_supplement=$data->sheets[0]['cells'][$i][25];//餐贴
$bonus=$data->sheets[0]['cells'][$i][26];//奖金
$commission=$data->sheets[0]['cells'][$i][27];//提成
$additional=$data->sheets[0]['cells'][$i][28];//正补项
$negative_buckle=$data->sheets[0]['cells'][$i][29];//负扣项
$should_pay=$data->sheets[0]['cells'][$i][30];//应发工资
$social_security_withholding=$data->sheets[0]['cells'][$i][31];//社保代扣
$fund=$data->sheets[0]['cells'][$i][32];//公积金代扣
$salary_payable=$data->sheets[0]['cells'][$i][33];//应付工资
$individual_income_tax=$data->sheets[0]['cells'][$i][34];//个人所得税
$real_wages=$data->sheets[0]['cells'][$i][35];//实发工资
$secrecy_salary=$data->sheets[0]['cells'][$i][36];//保密工资
$provident_fund_subsidies=$data->sheets[0]['cells'][$i][37];//公积金补贴
$clock_real_hair=$data->sheets[0]['cells'][$i][38];//打卡实发
$rs=explode("@",$list_email);
$user_name = $rs['0'];
$body = '
<html>
<head>
<title>my test email</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body bgcolor="#FFFFFF" >
<table width=100% border="2" cellpadding="3" cellspacing="0" bordercolor="#808080"> 
<tr bgcolor="#84A9E1"> 
<td align="center">公司名称</td> 
<td align="center">一级部门</td> 
<td align="center">工号</td> 
<td align="center">员工姓名</td> 
<td align="center">应出勤</td> 
<td align="center">实出勤</td> 
<td align="center">出差</td> 
<td align="center">事假</td> 
<td align="center">病假</td> 
<td align="center">婚假</td> 
<td align="center">产假</td> 
<td align="center">陪产假</td> 
<td align="center">丧假</td> 
<td align="center">年休假</td> 
<td align="center">法定节假日</td> 
<td align="center">法定节假日加班</td> 
<td align="center">周末加班</td> 
<td align="center">工作日加班</td> 
<td align="center">值班</td> 
<td align="center">旷工</td> 
<td align="center">工资标准</td>
<td align="center">出勤工资</td> 
<td align="center">加班工资</td> 
<td align="center">餐贴</td> 
<td align="center">奖金</td> 
<td align="center">提成</td> 
<td align="center">正补项</td> 
<td align="center">负扣项</td> 
<td align="center">应发工资</td> 
<td align="center">社保代扣</td> 
<td align="center">公积金代扣</td> 
<td align="center">应付工资</td> 
<td align="center">个人所得税</td> 
<td align="center">实发工资</td> 
<td align="center">保密工资</td> 
<td align="center">公积金补贴</td> 
<td align="center">打卡实发</td>  
</tr>
<td align="center">'.$company_name.'</td> 
<td align="center">'.$department.'</td> 
<td align="center">'.$job_num.'</td> 
<td align="center">'.$name.'</td> 
<td align="center">'.$work_allday.'</td> 
<td align="center">'.$work_realday.'</td> 
<td align="center">'.$business_travel.'</td> 
<td align="center">'.$private_leave.'</td> 
<td align="center">'.$sick_leave.'</td> 
<td align="center">'.$marital_leave.'</td>
<td align="center">'.$maternity_leave.'</td> 
<td align="center">'.$maternity_leave_company.'</td> 
<td align="center">'.$jbereavement_leave.'</td> 
<td align="center">'.$annual_leave.'</td> 
<td align="center">'.$statutory_holidays.'</td> 
<td align="center">'.$statutory_holiday_overtime.'</td> 
<td align="center">'.$weekend_overtime.'</td> 
<td align="center">'.$working_overtime.'</td> 
<td align="center">'.$be_on_duty.'</td> 
<td align="center">'.$neglect_work.'</td> 
<td align="center">'.$wage.'</td> 
<td align="center">'.$attendance_standards.'</td> 
<td align="center">'.$overtime_wages.'</td> 
<td align="center">'.$meal_supplement.'</td> 
<td align="center">'.$bonus.'</td> 
<td align="center">'.$commission.'</td> 
<td align="center">'.$additional.'</td> 
<td align="center">'.$negative_buckle.'</td> 
<td align="center">'.$should_pay.'</td> 
<td align="center">'.$social_security_withholding.'</td> 
<td align="center">'.$fund.'</td> 
<td align="center">'.$salary_payable.'</td> 
<td align="center">'.$individual_income_tax.'</td> 
<td align="center" style="color:red;">'.$real_wages.'</td> 
<td align="center">'.$secrecy_salary.'</td> 
<td align="center">'.$provident_fund_subsidies.'</td> 
<td align="center" style="color:red;">'.$clock_real_hair.'</td> 
</table>
</body>
</html>
		';
smtp_mail($list_email, '融都科技工资单', $body, 'www.erongdu.com', $user_name);
redirect("?action=$i"); 
}
echo "<div class='waiting'>正在发送邮件,请稍等......</div>";
}else {
if($i>=1){
 echo "<div class='sucess'>{$i}封邮件已经全部发送完毕</div>";
 }
 exit;
  } 

?>
