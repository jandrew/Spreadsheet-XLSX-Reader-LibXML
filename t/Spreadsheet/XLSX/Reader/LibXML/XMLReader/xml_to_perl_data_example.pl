#!/usr/bin/env perl
$|=1;
use Data::Dumper;
use	MooseX::ShortCut::BuildInstance qw( build_instance );
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader;
use	Spreadsheet::XLSX::Reader::LibXML::Error;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData;
my  $test_file = '../../../../../test_files/xl/sharedStrings.xml';
my  $test_instance	=	build_instance(
		package => 'TestIntance',
		superclasses =>[ 'Spreadsheet::XLSX::Reader::LibXML::XMLReader', ],
		add_roles_in_sequence =>[ 'Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData', ],
		file => $test_file,
		error_inst => Spreadsheet::XLSX::Reader::LibXML::Error->new,
	);
map{ $test_instance->next_element( 'si' ) }( 0..15 );# Go somewhere interesting
print Dumper( $test_instance->parse_element ) . "\n";