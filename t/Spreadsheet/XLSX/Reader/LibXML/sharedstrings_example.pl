#!/usr/bin/env perl
$|=1;
use Data::Dumper;
use MooseX::ShortCut::BuildInstance qw( build_instance );
use lib '../../../../../../lib';
use Spreadsheet::XLSX::Reader::LibXML::Error;
use Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings;

my $file_instance = build_instance(
      package      => 'SharedStringsInstance',
      superclasses => ['Spreadsheet::XLSX::Reader::LibXML::XMLReader::SharedStrings'],
      file         => '../../../../../test_files/xl/sharedStrings.xml',
      error_inst   => Spreadsheet::XLSX::Reader::LibXML::Error->new,
	);
print Dumper( $file_instance->get_shared_string_position( 3 ) );
print Dumper( $file_instance->get_shared_string_position( 12 ) );