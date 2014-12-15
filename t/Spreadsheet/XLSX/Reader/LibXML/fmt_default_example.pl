#!/usr/bin/env perl
package MyPackage;
use Moose;
use lib '../../../../../lib';
with 'Spreadsheet::XLSX::Reader::LibXML::FmtDefault';

package main;

my $parser = MyPackage->new;
print '(' . join( ', ', $parser->get_defined_excel_format( 14 ) ) . ")\n";