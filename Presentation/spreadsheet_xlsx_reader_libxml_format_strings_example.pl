#!/usr/bin/env perl
 
use strict;
use warnings;
use DateTimeX::Format::Excel;
use DateTime::Format::Flexible;
use Type::Tiny;
use Types::Standard qw( Num Str );
use lib '../lib';
use Spreadsheet::XLSX::Reader::LibXML;
$| = 1;# To show where the undefs occur
my $workbook =	Spreadsheet::XLSX::Reader::LibXML->new(
					file_name => '../t/test_files/TestBook.xlsx'
				);
 
if ( !$workbook->has_file_name ) {
    die $workbook->error(), ".\n";
}

# Uglyness here
my $system_lookup = {
		'1900' => 'win_excel',
		'1904' => 'apple_excel',
	};
my	@args_list	= ( system_type => $system_lookup->{$workbook->get_epoch_year} );
my	$converter	= DateTimeX::Format::Excel->new( @args_list );
my	$string_via	= sub{ 
						my	$str = $_[0];
						return DateTime::Format::Flexible->parse_datetime( $str );#my	$dt	= 
						#~ return $dt->format_cldr( 'yyyy-M-d' );
					};
my	$num_via	= sub{
						my	$num = $_[0];
						return $converter->parse_datetime( $num );#my	$dt = 
						#~ return $dt->format_cldr( 'yyyy-M-d' );
					};
my	$date_time_from_value = Type::Coercion->new(
		type_coercion_map => [ Num, $num_via, Str, $string_via, ],
	);
my	$date_time_type = Type::Tiny->new(
		name		=> 'Custom_date_type',
		constraint	=> sub{ ref($_) eq 'DateTime' },
		coercion	=> $date_time_from_value,
	);
my	$string_type = Type::Tiny->new(
		name		=> 'YYYYMMDD',
		constraint	=> sub{
			!$_ or (
			$_ =~ /^\d{4}\-(\d{2})-(\d{2})$/ and
			$1 > 0 and $1 < 13 and $2 > 0 and $2 < 32 )
		},
		coercion	=> Type::Coercion->new(
			type_coercion_map =>[
				$date_time_type->coercibles, sub{
					my $tmp = $date_time_type->coerce( $_ );
					$tmp->format_cldr( 'yyyy-MM-dd' ) 
				},
			],
		),
);

for my $worksheet ( $workbook->worksheets() ) {
	
	print $worksheet->get_name . "\n";# Not in the SYNOPSIS ( ParseExcel uses get_name )
	next if $worksheet->get_name ne 'Sheet1';# Not in the SYNOPSIS
	$worksheet->set_custom_formats( {
		E10	=> $date_time_type,# Incorperates the windows vs apple on the fly!
		10	=> $string_type,
		D14	=> $workbook->parse_excel_format_string( 'yyyy mmmm d, h:mm AM/PM' ),#Money shot!
	} );
    my ( $row_min, $row_max ) = $worksheet->row_range();
    my ( $col_min, $col_max ) = $worksheet->col_range();
 
    for my $row ( $row_min .. $row_max ) {
        for my $col ( $col_min .. $col_max ) {
 
            my $cell = $worksheet->get_cell( $row, $col );
            next unless $cell;
 
            print "Row, Col    = ($row, $col)\n";
			print "CellID      = " . $cell->cell_id . "\n";
            print "Value       = " . ($cell->value()//'undef') . "\n";# $cell->value()
            print "Unformatted = " . ($cell->unformatted()//'undef') . "\n";# $cell->unformatted()
            print "\n";
        }
    }
}