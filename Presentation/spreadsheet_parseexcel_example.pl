#!/usr/bin/env perl
 
use strict;
use warnings;
use Spreadsheet::ParseExcel;
 
my $parser   = Spreadsheet::ParseExcel->new();
my $workbook = $parser->parse( '../t/test_files/TestBook.xls' );
 
if ( !defined $workbook ) {
    die $parser->error(), ".\n";
}
 
for my $worksheet ( $workbook->worksheets() ) {
	
	print $worksheet->get_name . "\n";# Not in the SYNOPSIS
	next if $worksheet->get_name ne 'Sheet1';# Not in the SYNOPSIS
	
    my ( $row_min, $row_max ) = $worksheet->row_range();
    my ( $col_min, $col_max ) = $worksheet->col_range();
 
    for my $row ( $row_min .. $row_max ) {
        for my $col ( $col_min .. $col_max ) {
 
            my $cell = $worksheet->get_cell( $row, $col );
            next unless $cell;
 
            print "Row, Col    = ($row, $col)\n";
            print "Value       = ", $cell->value(),       "\n";
            print "Unformatted = ", $cell->unformatted(), "\n";
            print "\n";
        }
    }
}