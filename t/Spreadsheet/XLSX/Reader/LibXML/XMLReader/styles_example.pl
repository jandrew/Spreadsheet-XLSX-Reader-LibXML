#!/usr/bin/env perl
$|=1;
use Data::Dumper;
use MooseX::ShortCut::BuildInstance qw( build_instance );
use lib '../../../../../../lib';
use Spreadsheet::XLSX::Reader::LibXML::Error;
use Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles;
use Spreadsheet::XLSX::Reader::LibXML::FmtDefault;

my $file_instance = build_instance(
      package      => 'StylesInstance',
      superclasses => ['Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles'],
      file         => '../../../../../test_files/xl/styles.xml',
      error_inst   => Spreadsheet::XLSX::Reader::LibXML::Error->new,
	  format_inst  => Spreadsheet::XLSX::Reader::LibXML::FmtDefault->new(
						epoch_year	=> 1904,
						error_inst	=> Spreadsheet::XLSX::Reader::LibXML::Error->new,
					),
	);
print Dumper( $file_instance->get_format_position( 2 ) );