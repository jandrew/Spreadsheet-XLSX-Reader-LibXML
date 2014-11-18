#!/usr/bin/env perl
use strict;
use warnings;
use lib '../../../../lib';
use Spreadsheet::XLSX::Reader::LibXML;

my $parser   = Spreadsheet::XLSX::Reader::LibXML->new();
my $workbook = $parser->parse( '../../../test_files/TestBook.xlsx' );

if ( !defined $workbook ) {
	die $parser->error(), "\n";
}

for my $worksheet ( $workbook->worksheets() ) {

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
	last;# In order not to read all sheets
}