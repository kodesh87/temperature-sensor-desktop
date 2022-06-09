<?php
	//local
	$mysql_host = "localhost";
	$mysql_database = "data";
	$mysql_user = "root";
	$mysql_password = "";
/*	
	//testing.hostzi
	$mysql_host = "mysql11.000webhost.com";
	$mysql_database = "a2060799_db";
	$mysql_user = "a2060799_root";
	$mysql_password = "root123";
	
	//terrysitumorang2.tk
	$mysql_host = "mysql11.000webhost.com";
	$mysql_database = "a8994901_db";
	$mysql_user = "a8994901_root";
	$mysql_password = "root123";
	
	//rumahweb
	$mysql_host = "localhost";
	$mysql_database = "wirasi89_db";
	$mysql_user = "wirasi89_root";
	$mysql_password = "101187291210wira";
*/
	$connection = mysql_connect("$mysql_host","$mysql_user","$mysql_password");
	
	mysql_select_db($mysql_database,$connection);
?>