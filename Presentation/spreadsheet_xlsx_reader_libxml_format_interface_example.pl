#!/usr/bin/env perl
use lib '../lib',
		'../../Log-Shiras/lib';
use MooseX::ShortCut::BuildInstance 'build_instance';
#~ use Log::Shiras::Switchboard qw( :debug );
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(#
###LogSD						name_space_bounds =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'trace',
###LogSD							},
###LogSD						},
###LogSD						reports =>{
###LogSD							log_file =>[ Print::Log->new ],
###LogSD						},
###LogSD					);
###LogSD	use Log::Shiras::Telephone;
###LogSD	my $phone = Log::Shiras::Telephone->new;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::FmtDefault;
use Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings;
use	Spreadsheet::XLSX::Reader::LibXML::FormatInterface;
use	Spreadsheet::XLSX::Reader::LibXML;
my	$formatter = build_instance(
					package => 'FormatInstance',
					superclasses => [ 'Spreadsheet::XLSX::Reader::LibXML::FmtDefault' ],# Inject your customized format class here
					add_roles_in_sequence =>[qw(
						Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings
						Spreadsheet::XLSX::Reader::LibXML::FormatInterface
					)],
					target_encoding => 'latin1',# Adjust only the output encoding here
					datetime_dates	=> 1,
				);
	$formatter->set_defined_excel_formats( 0x2C => 'MyCoolFormatHere' ); #Set specific default custom formats here
my	$parser	= Spreadsheet::XLSX::Reader::LibXML->new;
print "$parser\n";
my	$workbook = $parser->parse( '../t/test_files/TestBook.xlsx', $formatter );
print "$workbook\n";
	$workbook = Spreadsheet::XLSX::Reader::LibXML->new(# This is an alternate way
					file_name		=> '../t/test_files/TestBook.xlsx',
					formatter_inst	=> $formatter,
				);
print "$workbook\n";

###LogSD	package Print::Log;
###LogSD	use Data::Dumper;
###LogSD	sub new{
###LogSD		bless {}, shift;
###LogSD	}
###LogSD	sub add_line{
###LogSD		shift;
###LogSD		my @input = ( ref $_[0]->{message} eq 'ARRAY' ) ? 
###LogSD						@{$_[0]->{message}} : $_[0]->{message};
###LogSD		my ( @print_list, @initial_list );
###LogSD		no warnings 'uninitialized';
###LogSD		for my $value ( @input ){
###LogSD			push @initial_list, (( ref $value ) ? Dumper( $value ) : $value );
###LogSD		}
###LogSD		for my $line ( @initial_list ){
###LogSD			$line =~ s/\n$//;
###LogSD			$line =~ s/\n/\n\t\t/g;
###LogSD			push @print_list, $line;
###LogSD		}
###LogSD		printf( "| level - %-6s | name_space - %-s\n| line  - %04d   | file_name  - %-s\n\t:(\t%s ):\n", 
###LogSD					$_[0]->{level}, $_[0]->{name_space},
###LogSD					$_[0]->{line}, $_[0]->{filename},
###LogSD					join( "\n\t\t", @print_list ) 	);
###LogSD		use warnings 'uninitialized';
###LogSD	}

###LogSD	1;