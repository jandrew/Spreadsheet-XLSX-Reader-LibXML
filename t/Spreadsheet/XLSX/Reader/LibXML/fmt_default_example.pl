#!/usr/bin/env perl
package MyPackage;
use Moose;
use lib '../../../../../lib';
with 'Spreadsheet::XLSX::Reader::LibXML::FmtDefault';

sub get_log_space{}

package main;

my $parser = MyPackage->new;
print '(' . join( ', ', $parser->get_defined_excel_format( 14 ) ) . ")\n";