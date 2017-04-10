<!DOCTYPE html>
<html xmls="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" charset="text/html utf-8">
	<title>Alunni</title>
</head>
<body>
<?php
	//percorso database
	$percorso = realpath("db.mdb");
	//Stringa di connesione ADO DB
	$sc = "Driver={Microsoft Access Driver (*.mdb)}; Dbq=".$percorso.";";
	//Crea due oggetti COM contententi gli oggetti Connection e Recordset
	$cn = new COM("ADODB.Connection") or die("Non va ADO");
	$rs = new COM("ADODB.Recordset");
	//Apro la Connection ed il Recordset
	$cn->Open($sc);
	$rs->Open("SELECT * FROM alunni", $cn);
	//Stampa tabella con riga di intestazione
	echo "<table><tr><th>ID</th><th>Nome</th><th>Cognome</th></tr>";
	$alt = false;
	//Ciclo di lettura e stampa recordset
	while (!$rs->eof)
	{
		//In base al contenuto della variabile $alt viene inserito lo stile altClass
		//per ottenere l'effetto del colore di sfondo alternato dalle righe
		$altClass = $alt ? " class='alt'" : "";
		echo "<tr>";
		//Stampa di ogni singolo campo
		echo "<td".$altClass.">".$rs->fields['ID']."</td>";
		echo "<td".$altClass.">".$rs->fields['Nome']."</td>";
		echo "<td".$altClass.">".$rs->fields['Cognome']."</td>";
		echo "</tr>";
		//Metodo movenext() per passare al record succcessivo
		$rs->movenext();
		//negazione del contenuto di $alt
		$alt = !$alt;
	}
	echo "</table";
	//Chiusura connesione
	$cn->close();
?>
</body>
</html>