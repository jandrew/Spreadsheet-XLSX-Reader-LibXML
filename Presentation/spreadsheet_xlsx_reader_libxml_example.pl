#!/usr/bin/env perl
$| = 1;# To show where the undefs occur
 
use strict;
use warnings;
use lib '../lib';
use Spreadsheet::XLSX::Reader::LibXML;
my $parser =	Spreadsheet::XLSX::Reader::LibXML->new();
my $workbook = $parser->parse( '../t/test_files/TestBook.xlsx' );
 
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
            print "Value       = " . ($cell->value()//'undef') . "\n";# $cell->value()
            print "Unformatted = " . ($cell->unformatted()//'undef') . "\n";# $cell->unformatted()
            print "\n";
        }
    }
}
$workbook = undef;
$parser = undef;