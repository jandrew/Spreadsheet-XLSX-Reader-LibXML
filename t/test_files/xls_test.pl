#! perl

use strict;
use Spreadsheet::ParseExcel;
use Data::Dumper;

my $parser   = Spreadsheet::ParseExcel->new();
my $workbook = $parser->parse('TestBook.xls');

if ( !defined $workbook ) {
	die $parser->error(), ".\n";
}

my $worksheet = $workbook->worksheet( 'Sheet1' );
my $cell = $worksheet->get_cell( 4, 0 );
print	$cell->value . "\n";
print	Dumper( $cell->get_rich_text );

$workbook = $parser->parse('Rich.xls');

if ( !defined $workbook ) {
	die $parser->error(), ".\n";
}

$worksheet = $workbook->worksheet( 'Sheet1' );
$cell = $worksheet->get_cell( 1, 0 );
print	$cell->value . "\n";
print	Dumper( $cell->get_rich_text );