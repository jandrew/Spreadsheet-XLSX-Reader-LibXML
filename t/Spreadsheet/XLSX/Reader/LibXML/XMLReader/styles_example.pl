#!/usr/bin/env perl
$|=1;
use Data::Dumper;
use MooseX::ShortCut::BuildInstance qw( build_instance );
use lib '../../../../../../lib';
use Spreadsheet::XLSX::Reader::LibXML::Error;
use Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles;

my $file_instance = build_instance(
      package      => 'StylesInstance',
      superclasses => ['Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles'],
      file         => '../../../../../test_files/xl/styles.xml',
      error_inst   => Spreadsheet::XLSX::Reader::LibXML::Error->new,
      add_roles_in_sequence => [qw(
         Spreadsheet::XLSX::Reader::LibXML::FmtDefault
         Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings
      )],
	);
print Dumper( $file_instance->get_format_position( 2 ) );