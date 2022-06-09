<?php
	include "config.php";

	foreach ($_GET as $name => $var) {
		$temp = $var;
	}

	$query = "INSERT INTO history VALUES('','$temp',CURRENT_TIMESTAMP)";
	$return = mysql_query($query);
?>