<?php

error_reporting(E_ALL); // for testing, expose all errors

require_once "Sheets.php";
require_once "ESELSearch.php";
require_once "Spreadsheet_Excel_Reader.php";

$filename = isset($_GET['esel']) ? $_GET['esel'] : '53P_Revision_1_ESEL.xls';

// must be .xls format
$filename = "ESELs/$filename";

$esel = new ESELSearch($filename);
$esel->extract_esel_data();


?>
<h1>ESEL Wikifier</h1>
<h2>Instructions</h2>
<p>This program is a work in progress. At present you can choose from the following ESELs to test:</p>
<ul>
<?php
$files = array(
	"35S_ESEL.xls",
	"48P_ESEL_Rev_1.xls",
	"49P_ESEL_Rev_1.xls",
	"52P_Baseline_ESEL.xls",
	"53P_Revision_1_ESEL.xls",
	"53P_Revision_1_ESEL.xlsx",
);
foreach($files as $file) {
	echo "<li><a href='?esel=$file'>$file wikification</a> - <a href='ESELs/$file'>View ESEL</a></li>";
}
?>
</ul>
<p>Note that the last one, the XLSX file, <strong>WILL NOT WORK</strong>. Also note that not all ESELs have all of the sections (Tools up, EMU down, Tools disposed, etc), which causes some error messages on this page. The output still appears to function, though.</p>
<h2>Wikitext Table Output</h2><?php
echo "<pre>";
print_r($esel->wikify_esel_data());

echo "</pre>";

// $sheets = $esel->comparator();

// $count = 0;
// $tabs = "";
// $tab_contents = "";
// foreach($sheets as $sheet_name => $sheet) {

	// // HERE YOU'D NEED TO LOOP THROUGH THE SHEETS AND DISPLAY YOUR CONTENT
	
// }
