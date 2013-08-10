<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<title>Films API</title>
<style>
html,*{
	box-sizing:border-box;
}
body{
	padding:15px;
	margin:0;
	background-color: black;
}
ol{
	margin:0;
	padding:0;
}
li{
	position:relative;
	display:block;
	float:left;
	width:175px;
	height:260px;
	padding:0;
	margin-right:15px;
	margin-bottom:15px;
	overflow:hidden;
	color:white;
	background-color:black;
	border: #000 solid 2px;
	-webkit-transition: all .2s;
}
li:hover{
	border: #FF6 solid 2px;
}
li .info{
	position: absolute;
	padding:15px;
	z-index:2;
	opacity:0;
	-webkit-transition: opacity .5s;
}
li:hover .info{
	opacity:1;
	-webkit-transition: opacity .3s;
}
li a{
	color: white;
}
li img{
	position: absolute;
	z-index:1;
	opacity:1;
	-webkit-transition: opacity .5s;
	
}
li:hover img{
	position: absolute;
	opacity:0.2;
	-webkit-transition: opacity .3s;
}

</style>
</head>
<body>
<iframe width="560" height="315" src="http://www.youtube.com/embed/videoseries?list=PL5839BD8A1A615C59" frameborder="0" allowfullscreen></iframe>
<br>
<?php
// ID y Clave secreta de la aplicación creada en http://series.ly/scripts/info/dev.php
$app_id = 1304;
$secret_key = "nVkynuenxZhXbynYbDne";

// Obtener auth_token
$auth_token = file_get_contents ("http://series.ly/api/auth.php?api=".$app_id."&secret=".$secret_key);

$search = urlencode($_GET["s"]);
$url = "http://series.ly/api/search.php?auth_token=".$auth_token."&search=".$search."&type=film&format=xml";
?>

<p><a href="<?php echo $url ?>">Buscar <?= $_GET["s"] ?> en Series.ly</a></p>
<?
function getXML($file){
	$xml_file = file_get_contents ($file);
	
	if (empty($xml_file)) die("Error: No se ha podido conectar.");
	
	preg_match_all("|<item>(.*)</item>|sU", $xml_file, $items);
	$nodes = array();

	$total =0;
	foreach ($items[1] as $key => $item) {
		preg_match("|<title>(.*)</title>|s", $item, $titulo);
		preg_match("|<year>(.*)</year>|s", $item, $year);
		preg_match("|<idFilm>(.*)</idFilm>|s", $item, $enlace);
		preg_match("|<poster>(.*)</poster>|s", $item, $poster);

		$nodes[$key]['title'] = $titulo[1];
		$nodes[$key]['year'] = $year[1];
		$nodes[$key]['idFilm'] = $enlace[1];
		$nodes[$key]['poster'] = $poster[1];
		
		$total += 1;
	}
	echo "<ol>";
	for ($i=0; $i<$total; $i++) { ?>
<li>
        <img src="<?= $nodes[$i]['poster'] ?>">
        <div class="info">
        <h3><?= $nodes[$i]['title'] ?> <?= $nodes[$i]['year'] ?></h3>
		<a href="http://series.ly/pelis/peli-<?= $nodes[$i]['idFilm'] ?>" target="_blank">Series.ly</a><br>
		<a href="http://thepiratebay.se/search/<?= $nodes[$i]['title'] ?>" target="_self">The Pirate Bay </a>
        </div>
        </li>
    <?php }
	echo "</ol>";
	$xml_file = "";
	}
	?>



<?php getXML($url); ?>
</body>
</html>
