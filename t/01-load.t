#!/usr/bin/env perl
### Test that the module(s) load!(s)
use	Test::More;
BEGIN{ use_ok( Test::Pod, qw( 1.48 ) ) };
BEGIN{ use_ok( TAP::Formatter::Console ) };
BEGIN{ use_ok( TAP::Harness ) };
BEGIN{ use_ok( TAP::Parser::Aggregator ) };
BEGIN{ use_ok( version ) };
BEGIN{ use_ok( Test::Moose ) };
BEGIN{ use_ok( Data::Dumper ) };
BEGIN{ use_ok( Capture::Tiny, qw( capture_stderr capture_stdout ) ) };
BEGIN{ use_ok( Carp, qw( cluck ) ) };
BEGIN{ use_ok( Clone, qw( clone ) ) };
BEGIN{ use_ok( XML::LibXML::Reader ) };
BEGIN{ use_ok( Type::Tiny, 1.000 ) };
BEGIN{ use_ok( Moose ) };
BEGIN{ use_ok( MooseX::StrictConstructor ) };
BEGIN{ use_ok( MooseX::HasDefaults::RO ) };
BEGIN{ use_ok( Archive::Zip ) };
BEGIN{ use_ok( File::Temp ) };
BEGIN{ use_ok( DateTimeX::Format::Excel, 0.012 ) };
BEGIN{ use_ok( MooseX::ShortCut::BuildInstance, 1.026 ) };
BEGIN{ use_ok( MooseX::ShortCut::BuildInstance, qw( build_instance ) ) };
BEGIN{ use_ok( DateTime::Format::Flexible ) };
use	lib '../lib', 'lib';
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::Cell, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::CellToColumnRow, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::Chartsheet, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::Error, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::FmtDefault, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::FormatInterface, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::Row, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::SharedStrings, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::Styles, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::Types, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::WorkbookFileInterface, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::WorkbookMetaInterface, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::WorkbookPropsInterface, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::WorkbookRelsInterface, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::Worksheet, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::WorksheetToRow, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::XMLReader, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::XMLReader::ExtractFile, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::XMLReader::NamedStyles, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::XMLReader::PositionStyles, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::XMLToPerlData, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::ZipReader, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML::ZipReader::ExtractFile, 0.044 ) };
BEGIN{ use_ok( Spreadsheet::XLSX::Reader::LibXML, 0.044 ) };
done_testing();