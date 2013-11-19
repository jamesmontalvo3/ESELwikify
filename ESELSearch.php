<?php

class ESELSearch {

	public $excel_data;
	public $columnsNumbers;
	public $results;
	
	public function __construct ($filepath) {
		$this->excel_data = new Spreadsheet_Excel_Reader($filepath, false);
		$this->get_row_ranges();
		$this->determine_columns();
		$this->generate_column_finder();
	}
	
	protected function get_row_ranges () {
		global $egRangeIdentifiers, $egRangeIdentifierColumn;
		
		$row = 1;
		$count = count($egRangeIdentifiers);
		
		$prev = false;
		for($i=0; $i<$count; $i++) {
			$initialRow = $row;
		
			// find starting point
			while( ! isset($egRangeIdentifiers[$i]['start']) ) {
				if ($this->excel_data->val($row,$egRangeIdentifierColumn,0) == $egRangeIdentifiers[$i]['label']) {
					$egRangeIdentifiers[$i]['start'] = $row+1;
				
					if ( $prev !== false )
						$egRangeIdentifiers[$prev]['end'] = $row-1;
					
					$found = true;
					break; // not really necessary because of while condition...
				}

				$row++;
					
				if ($row > 1000) {
					$row = $initialRow; // this label must not exist. Skip, reset row to initial.
					$found = false;
					break;
				}
			}
		
			if ($found)
				$prev = $i;
		}

		// above for-loop with nested while-loop found start of all sections
		// and end of all except the last. Find the last end
		$numBlank = 0;
		while ( $numBlank < 11 ) {
			$colVal = trim( $this->excel_data->val($row,$egRangeIdentifierColumn,0) );
			if ( $colVal == '' ) {
				if($numBlank == 0)
					$egRangeIdentifiers[$prev]['end'] = $row;
				$numBlank++;
			}
			else if ($numBlank > 0)
				$numBlank = 0; // reset to zero if 
			
			$row++;
		}
		
		return true;
	}
	
	protected function determine_columns () {
		for($row=1; $row<15; $row++) {
			if ($this->excel_data->val($row,5,0) == "Part Number")
				return true;
			else if ($this->excel_data->val($row,5,0) == "Description") {
				$this->shift_columns_for_soyuz();
				return true;
			}
		}
		// if did not return some point in the for loop, P/N and description columns never found
		throw new Exception('Could not find "Part Number" or "Description" columns in expected locations.');
		return false;
	}
	
	public function shift_columns_for_soyuz () {
		global $egColumns;
		$newColumns = array();
		foreach($egColumns as $colNum => $colName) {
			if ($colNum == 4)
				$newColumns[4] = "Line Item 2";
			if ($colNum > 3)
				$newColumns[$colNum+1] = $colName;
			else
				$newColumns[$colNum] = $colName;
		}
		$egColumns = $newColumns;
	}
	
	public function generate_column_finder () {
		global $egColumns;
		
		$columnsNumbers = array();
		foreach ($egColumns as $colnum => $colname) {
			$columnsNumbers[$colname] = $colnum;
		}
		$this->columnsNumbers = $columnsNumbers;
	}
	
	public function get_col_num ($colname) {
		return $this->columnsNumbers[$colname];
	}
	
	public function extract_esel_data () {
		global $egColumns, $egSheetNum, $egRangeIdentifiers;
		
		$results = array();
		$count = 0; // since $row won't start at zero, count with this.
		
		foreach($egRangeIdentifiers as $rids) {
			
			if ( $rids['id'] == 'up-tools' || $rids['id'] == 'up-emu' )
				$updown = 'up';
			else if ( $rids['id'] == 'down-tools' || $rids['id'] == 'down-emu' )
				$updown = 'down';
			else if ( $rids['id'] == 'dispose-tools' || $rids['id'] == 'dispose-emu' )
				$updown = 'dispose';
			else
				die( "something went wrong 3451" ); // FIXME: handle this properly
			
			if ( isset($rids['start']) && isset($rids['end']) ) {
				for($row=$rids['start']; $row<=$rids['end']; $row++) {
					
					if ( $this->row_has_qty_value($row) ) {
						foreach($egColumns as $colNum => $colName) {
							$results[$count][$colName] = $this->excel_data->val($row,$colNum,$egSheetNum);
						}
						$results[$count]['updown'] = $updown;
						$count++;
					}
				}
			}
		}
		
		return $this->results = $results;
	
	}
	
	public function row_has_qty_value ( $row ) {
		global $egSheetNum;
		
		$upqty   = $this->excel_data->val($row, $this->get_col_num("Qty Up"), $egSheetNum);	
		$downqty = $this->excel_data->val($row, $this->get_col_num("Qty Down"), $egSheetNum);
		
		return ($upqty > 0 || $downqty > 0);
	}
	
	public function wikify_esel_data () {
	
		$out = "{| class=\"wikitable smwtable sortable\"\n";
		foreach( $this->results[0] as $colname => $colval ){
			$out .= "! $colname\n";
		}
			
		foreach($this->results as $row => $cols) {
			$out .= "|-\n";
			foreach($cols as $colname => $colvalue) {
				$value = $colvalue;
				switch ($colname) {
					case "Serial Number":
						$value = implode(", ", preg_split("/[\s]+/",$colvalue)); // str_replace???
						break;
					case "Description":
						$value = "[[" . preg_replace("/[\s]+/", " ", $colvalue) . "]]";
						break;
					default;
						$value = $colvalue;
				}
				
				$out .= "| $value\n";
			
			}
			// $out .= "| [[" . $cols[4] . "]]\n";
			// $out .= "| " . $cols[5] . "\n";
			// $out .= "| " . ($cols[9] ? $cols[9] . " Up" : $cols[10] . " Down") . "\n";
			// $out .= "| " . $cols[20] . "\n";
		}
	
		return $out . "|}";
	}
	
}
