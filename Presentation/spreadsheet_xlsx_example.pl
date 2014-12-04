#!/usr/bin/env perl
 
use strict;
use warnings;
#~ use Data::Dumper;

use Spreadsheet::XLSX;

my $excel = Spreadsheet::XLSX->new('../t/test_files/TestBook.xlsx');# One step rather than two
 
#~ if ( !defined $workbook ) { # No error logging
    #~ die $parser->error(), ".\n";
#~ }
 
foreach my $sheet (@{$excel->{Worksheet}}) {# $workbook->worksheets()
 
	printf("Sheet: %s\n", $sheet->{Name});# $worksheet->get_name
	next if $sheet->{Name} ne 'Sheet1';# Not in the SYNOPSIS
	
	$sheet->{MaxRow} ||= $sheet->{MinRow};# $worksheet->row_range();

	foreach my $row ($sheet->{MinRow} .. $sheet->{MaxRow}){
	 
		$sheet->{MaxCol} ||= $sheet->{MinCol};# $worksheet->col_range();
		
		foreach my $col ($sheet->{MinCol} ..  $sheet->{MaxCol}){
                
			my $cell = $sheet->{Cells}[$row][$col];# $worksheet->get_cell( $row, $col );
			next unless $cell;
			 
			print "Row, Col    = ($row, $col)\n";
			print "Value       = " . $cell->{_Value} . "\n";# $cell->value()
			print "Unformatted = " . $cell->{Val} . "\n";# $cell->unformatted()
			print "\n";
			#~ print Dumper( $cell ) . "\n";
			
		}
		
	}
}