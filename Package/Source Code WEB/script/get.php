<?php
	include "config.php";

	$query = "SELECT temperature, time FROM history ORDER BY no DESC";
	$return = mysql_query($query);

	$data = mysql_fetch_row($return);
	
	echo ("$data[0] $data[1]");
?>