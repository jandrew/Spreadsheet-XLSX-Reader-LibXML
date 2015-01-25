#!/usr/bin/env perl
$|=1;
use Data::Dumper;
use lib '../../../../../../lib';
use Spreadsheet::XLSX::Reader::LibXML::Error;
use Spreadsheet::XLSX::Reader::LibXML::XMLReader::CalcChain;

my  $file_instance = Spreadsheet::XLSX::Reader::LibXML::XMLReader::CalcChain->new(
						file => '../../../../../test_files/xl/calcChain.xml',
						error_inst => Spreadsheet::XLSX::Reader::LibXML::Error->new,
					);
print Dumper( $file_instance->get_calc_chain_position( 2 ) );