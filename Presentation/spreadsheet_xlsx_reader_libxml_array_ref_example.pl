#!/usr/bin/env perl
$| = 1;
use strict;
use warnings;
use Data::Dumper;
use lib '../lib';
use Spreadsheet::XLSX::Reader::LibXML;
use ExampleTypes qw( DateTimeStringOneType );
$| = 1;# To show where the undefs occur
my $workbook =	Spreadsheet::XLSX::Reader::LibXML->new( #similar style to Spreadsheet::XLSX
					file_name => '../t/test_files/TestBook.xlsx',
					group_return_type => 'value',
				);

if ( !$workbook->has_file_name ) {
    die $workbook->error(), ".\n";
}

my	$worksheet = $workbook->worksheet( 'Sheet5' );
	$worksheet->set_custom_formats( {
		B2 => DateTimeStringOneType,
		C2 => DateTimeStringOneType,
		D2 => DateTimeStringOneType,
	} );
my $value;
while( !$value or $value ne 'EOF' ){
	$value = $worksheet->fetchrow_arrayref;
	print Dumper( $value );
}