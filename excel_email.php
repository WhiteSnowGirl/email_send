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
	    <div class="div2">�ϴ��ļ�</div>
	    <input name="file" type="file" class="inputstyle" id="file"/> 
	    
	    <div class="filename"></div>
	</div>
  	<input class="tijiao" type="submit" name="Submit" value="�ύ" />  
 
</form> 
</div>
</div>
<script>
$(document).ready(function(){
	$('.sucess').fadeOut(3000);
	$('.unread').fadeOut(2000);
	$('#file').change(function(){
		var name =$('#file').val();
		var arr=name.split('\\');//עsplit�������ַ����ַ����ָ� 
		var name=arr[arr.length-1];
		$('.filename').html(name);
	});
	
	
	// var arr=name.split('\\');//עsplit�������ַ����ַ����ָ� 
	// var my=arr[arr.length-1];//�����Ҫȡ�õ�ͼƬ���� 
	// alert(my);
});
</script>
</body>
</html>   
<?php
error_reporting(E_STRICT);
date_default_timezone_set('Asia/Shanghai');
// ���� PHPmailer�� �ļ�
require_once("class.phpmailer.php");
require_once 'reader.php';    
$data = new Spreadsheet_Excel_Reader();   
$data->setOutputEncoding('gbk');  
//����Email����
function smtp_mail ( $sendto_email, $subject, $body, $extra_hdrs, $user_name) {
$mail = new PHPMailer(); 
$mail->IsSMTP(); // send via SMTP 
$mail->Host = "smtp.163.com";   // SMTP servers 
$mail->SMTPAuth = true; // turn on SMTP authentication 
$mail->Username = "zzjing1224@163.com";  // SMTP username ע�⣺��ͨ�ʼ���֤����Ҫ�� @����
$mail->Password = "jing1224"; // SMTP password

$mail->From = "zzjing1224@163.com";  // ����������
$mail->FromName = ���������ܼ�; //   ������ ,����
$mail->CharSet = "GB2312";  // ����ָ���ַ�����
$mail->Encoding = "base64";

$mail->AddAddress($sendto_email,$user_name);// �ռ������������
$mail->AddReplyTo("","jingjing_test");

//$mail->WordWrap = 50; // set word wrap 
//$mail->AddAttachment("/var/tmp/file.tar.gz");// attachment  ����1
//$mail->AddAttachment("/tmp/image.jpg", "new.jpg"); //����2
$mail->IsHTML(true);   // send as HTML 
$mail->Subject = $subject;  
$mail->Body = $body;
$mail->AltBody ="text/html"; 
$date= date('Y-m-d H:i:s');
if($mail->Send()) 
{ 
   info_write("ok.txt","$user_name ���ͳɹ� $date");
} 
else {
   info_write("falied.txt","$user_name ʧ��,������Ϣ$mail->ErrorInfo  $date");
 }
}
// ����Email��������

// д�뷢�ͽ��������������־��¼
function info_write($filename,$info_log)
{
 $info.= $info_log;
 $info.="\r\n";
 $fp = fopen ($filename,a);
 fwrite($fp,$info);
 fclose($fp);
}

//��ʱ��תҳ�� �������� 1000��ʱ��,1��, �������Զ���
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

//��ȡ�ı� �ʼ���ַ  ��Ҳ���Զ� ���ݿ�
/*$filename = "email.txt";
$fp = fopen($filename,"r");
$contents = fread($fp,filesize($filename));
$list_email=explode("\r\n",$contents); 
$len=count($list_email);
fclose($fp);*/
// ����˵��(���͵�, �ʼ�����, �ʼ�����, ������Ϣ, �û���)
$i = $_GET['action']; 
if($_POST['Submit'])  
{ 
$data->read($_POST['file']);  
for ($i = 1; $i <= $data->sheets[0]['numRows']; $i++) {
$list_email=$data->sheets[0]['cells'][$i][1]; //�����ַ
$company_name=$data->sheets[0]['cells'][$i][2];//��˾����
$department=$data->sheets[0]['cells'][$i][3];//һ������
$job_num=$data->sheets[0]['cells'][$i][4];//����
$name=$data->sheets[0]['cells'][$i][5];//Ա������
$work_allday=$data->sheets[0]['cells'][$i][6];//Ӧ����
$work_realday=$data->sheets[0]['cells'][$i][7];//ʵ����
$business_travel=$data->sheets[0]['cells'][$i][8];//����
$private_leave =$data->sheets[0]['cells'][$i][9];//�¼�
$sick_leave=$data->sheets[0]['cells'][$i][10];//����
$marital_leave=$data->sheets[0]['cells'][$i][11];//���
$maternity_leave=$data->sheets[0]['cells'][$i][12];//����
$maternity_leave_company=$data->sheets[0]['cells'][$i][13];//�����
$jbereavement_leave=$data->sheets[0]['cells'][$i][14];//ɥ��
$annual_leave=$data->sheets[0]['cells'][$i][15];//���ݼ�
$statutory_holidays=$data->sheets[0]['cells'][$i][16];//�����ڼ���
$statutory_holiday_overtime=$data->sheets[0]['cells'][$i][17];//�����ڼ��ռӰ�
$weekend_overtime=$data->sheets[0]['cells'][$i][18];//��ĩ�Ӱ�
$working_overtime=$data->sheets[0]['cells'][$i][19];//�����ռӰ�
$be_on_duty=$data->sheets[0]['cells'][$i][20];//ֵ��
$neglect_work=$data->sheets[0]['cells'][$i][21];//����
$wage=$data->sheets[0]['cells'][$i][22];//���ʱ�׼
$attendance_standards=$data->sheets[0]['cells'][$i][23];//���ڹ���
$overtime_wages=$data->sheets[0]['cells'][$i][24];//�Ӱ๤��
$meal_supplement=$data->sheets[0]['cells'][$i][25];//����
$bonus=$data->sheets[0]['cells'][$i][26];//����
$commission=$data->sheets[0]['cells'][$i][27];//���
$additional=$data->sheets[0]['cells'][$i][28];//������
$negative_buckle=$data->sheets[0]['cells'][$i][29];//������
$should_pay=$data->sheets[0]['cells'][$i][30];//Ӧ������
$social_security_withholding=$data->sheets[0]['cells'][$i][31];//�籣����
$fund=$data->sheets[0]['cells'][$i][32];//���������
$salary_payable=$data->sheets[0]['cells'][$i][33];//Ӧ������
$individual_income_tax=$data->sheets[0]['cells'][$i][34];//��������˰
$real_wages=$data->sheets[0]['cells'][$i][35];//ʵ������
$secrecy_salary=$data->sheets[0]['cells'][$i][36];//���ܹ���
$provident_fund_subsidies=$data->sheets[0]['cells'][$i][37];//��������
$clock_real_hair=$data->sheets[0]['cells'][$i][38];//��ʵ��
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
<td align="center">��˾����</td> 
<td align="center">һ������</td> 
<td align="center">����</td> 
<td align="center">Ա������</td> 
<td align="center">Ӧ����</td> 
<td align="center">ʵ����</td> 
<td align="center">����</td> 
<td align="center">�¼�</td> 
<td align="center">����</td> 
<td align="center">���</td> 
<td align="center">����</td> 
<td align="center">�����</td> 
<td align="center">ɥ��</td> 
<td align="center">���ݼ�</td> 
<td align="center">�����ڼ���</td> 
<td align="center">�����ڼ��ռӰ�</td> 
<td align="center">��ĩ�Ӱ�</td> 
<td align="center">�����ռӰ�</td> 
<td align="center">ֵ��</td> 
<td align="center">����</td> 
<td align="center">���ʱ�׼</td>
<td align="center">���ڹ���</td> 
<td align="center">�Ӱ๤��</td> 
<td align="center">����</td> 
<td align="center">����</td> 
<td align="center">���</td> 
<td align="center">������</td> 
<td align="center">������</td> 
<td align="center">Ӧ������</td> 
<td align="center">�籣����</td> 
<td align="center">���������</td> 
<td align="center">Ӧ������</td> 
<td align="center">��������˰</td> 
<td align="center">ʵ������</td> 
<td align="center">���ܹ���</td> 
<td align="center">��������</td> 
<td align="center">��ʵ��</td>  
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
smtp_mail($list_email, '�ڶ��Ƽ����ʵ�', $body, 'www.erongdu.com', $user_name);
redirect("?action=$i"); 
}
echo "<div class='waiting'>���ڷ����ʼ�,���Ե�......</div>";
}else {
if($i>=1){
 echo "<div class='sucess'>{$i}���ʼ��Ѿ�ȫ���������</div>";
 }
 exit;
  } 

?>
