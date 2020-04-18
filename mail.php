<?
//******************************
//Mail.php for Adsense Alert
//Copyright 2005 AdseneAlert.com
//******************************

//Edit these options
$emailaddress="youremail@yoursite.com";
$mailfrom="youremail@yoursite.com";
$mailsubject="AdsenseAlert Update";

//Do not edit any of the files below



		$mailheaders="From: $mailfrom\n";
		$mailreslult = @mail($emailaddress, $mailsubject, $mailbody, $mailheaders);

		if(!$mailreslult)
		{
			print("<p>Failed to send email</p>");
		}
		else
		{
			print("<p>Email has been sent!</p>\n");
		}


?>
