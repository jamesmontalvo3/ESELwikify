<?php

/** 
 * Columns we care about. Column "A" = 1
 * 
 * This is the standard for Progress vehicles. Soyuz vehicles shift everything
 * from "description" one cell to the right and add another line item column.
 * This shift will be handled by the program.
 **/
$egColumns = array(
	2  => "Section",
	3  => "Line Item",
	4  => "Description", // item name
	5  => "Part Number",
	6  => "Serial Number",
	9  => "Qty Up",
	10 => "Qty Down",
	20 => "Notes", // AKA Remarks/Special Instructions
);

// should only be one sheet in Progress/Soyuz ESELs
// If extra sheets are added, the important sheet should be first (aka zero-th)
$egSheetNum = 0;

$egRangeIdentifierColumn = 2; // column where the labels below are found
$egRangeIdentifiers = array(
	array("id"=>"up-tools",      "label"=>"1. Launching Hardware Tools"),
	array("id"=>"up-emu",        "label"=>"1. Launching Hardware EMU"),
	array("id"=>"down-tools",    "label"=>"2. Return Hardware Tools"),
	array("id"=>"down-emu",      "label"=>"2. Return Hardware EMU"),
	array("id"=>"dispose-tools", "label"=>"3. Disposal  Hardware Tools"), // WTF? seriously, the extra space?
	array("id"=>"dispose-emu",   "label"=>"3. Disposal Hardware EMU"),
);