#!/usr/bin/env perl
$| = 1;
use strict;
use warnings;
use lib '../lib';
use Spreadsheet::XLSX::Reader::LibXML;
use ExampleTypes qw( DateTimeType DateTimeStringOneType DateTimeStringTwoType );
$| = 1;# To show where the undefs occur
my $workbook =	Spreadsheet::XLSX::Reader::LibXML->new( #similar style to Spreadsheet::XLSX
					file_name => '../t/test_files/TestBook.xlsx',
					group_return_type => 'unformatted',# 'value',
				);

if ( !$workbook->has_file_name ) {
    die $workbook->error(), ".\n";
}

my	$worksheet = $workbook->worksheet( 'Sheet1' );
	$worksheet->set_custom_formats( {
		E10	=> DateTimeType,
		10	=> DateTimeStringOneType,
		D14	=> DateTimeStringTwoType,
	} );
my $value;
while( !$value or $value ne 'EOF' ){
	$value = $worksheet->get_next_value;
	print( ($value//'undef') . "\n");
}