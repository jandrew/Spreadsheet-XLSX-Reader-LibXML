#########1 Test File for Spreadsheet::XLSX::Reader::LibXML::XMLDOM::Styles      8#########9
#!perl
BEGIN{ $ENV{PERL_TYPE_TINY_XS} = 0; }
$| = 1;

use	Test::Most tests => 35;
use	Test::Moose;
use	MooseX::ShortCut::BuildInstance v1.8 qw( build_instance );
use Data::Dumper;
use	lib
		'../../../../../../../Log-Shiras/lib',
		'../../../../../../lib',;
#~ use Log::Shiras::Switchboard qw( :debug );
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(#
###LogSD						name_space_bounds =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'trace',
###LogSD							},
###LogSD							Test =>{
###LogSD								_parse_the_file =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'warn',
###LogSD									},
###LogSD								},
###LogSD								_load_data_to_format =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'warn',
###LogSD									},
###LogSD								},
###LogSD								process_element_to_perl_data =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'warn',
###LogSD									},
###LogSD								},
###LogSD							},
###LogSD						},
###LogSD						reports =>{
###LogSD							log_file =>[ Print::Log->new ],
###LogSD						},
###LogSD					);
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::XMLDOM::Styles;
use	Spreadsheet::XLSX::Reader::LibXML::Error;
my	$test_file = ( @ARGV ) ? $ARGV[0] : '../../../../../test_files/xl/';
	$test_file .= 'styles.xml';
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'trace', message => [ "Test file is: $test_file" ] );
my  ( 
			$test_instance, $capture, $x, @answer, $error_instance, $format_instance,
	);
my 			$row = 0;
my 			@class_attributes = qw(
				file_name
				excel_region
				defined_excel_translations
				epoch_year
				error_inst
			);
my  		@class_methods = qw(
				new
				has_file_name
				get_excel_region
				get_epoch_year
				set_epoch_year
				set_error
				clear_error
				set_warnings
				if_warn
				get_numFmts
				get_cellXfs
				get_fonts
				get_numFmts
				get_cellStyleXfs
				get_tableStyles
				get_fills
				get_borders
				get_dxfs
				get_cellStyles
				encoding
				parse_excel_format_string
				get_format_position
				get_default_format_position
				process_element_to_perl_data
			);
				#~ get_number_format
###LogSD		$phone->talk( level => 'info', message => [ "easy questions ..." ] );
map{ 
has_attribute_ok
			'Spreadsheet::XLSX::Reader::LibXML::XMLDOM::Styles', $_,
										"Check that Spreadsheet::XLSX::Reader::LibXML::XMLDOM::Styles has the -$_- attribute"
} 			@class_attributes;
map{
can_ok		'Spreadsheet::XLSX::Reader::LibXML::XMLDOM::Styles', $_,
} 			@class_methods;

###LogSD		$phone->talk( level => 'info', message => [ "harder questions ..." ] );
lives_ok{
			$test_instance	=	Spreadsheet::XLSX::Reader::LibXML::XMLDOM::Styles->new(
									file_name	=> $test_file,
									log_space	=> 'Test',
									error_inst	=> Spreadsheet::XLSX::Reader::LibXML::Error->new(
										should_warn => 1,
										#~ should_warn => 0,# to turn off cluck when the error is set
									),
									epoch_year	=> 1904,
								);
}										"Prep a new Styles instance";

###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );
###LogSD		$phone->talk( level => 'info', message => [ "Number format:",  $test_instance->get_format_position( 2 ) ] );
ok			$format_instance = $test_instance->parse_excel_format_string( "[$-409]d\-mmm\-yy;@" ),
										"Create a number conversion from an excel format string";
#~ explain									Dumper( $format_instance );
is			$format_instance->( 37145 ), '12-Sep-05', #coercecoerce
										"... and see if it works";
is			$test_instance->get_format_position( 2, 'NumberFormat' )->name, 'DateTime164FromNum',
										"Check that the custom formatted number translation is named: DateTime164FromNum";
is			$test_instance->get_format_position( 7, 'font' )->{sz}, 14,
										"Check that number format position 7 has a font size set to: 14";
is			$test_instance->get_fonts->[ 1 ]->{color}->{rgb}, 'FFFF0000',
										"Check that font definition position 1 has the 'rgb' color set to: FFFF0000";
explain 								"...Test Done";
done_testing();

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
###LogSD			$line =~ s/\n/\n\t\t/g;
###LogSD			push @print_list, $line;
###LogSD		}
###LogSD		printf( "name_space - %-50s | level - %-6s |\nfile_name  - %-50s | line  - %04d   |\n\t:(\t%s ):\n", 
###LogSD					$_[0]->{name_space}, $_[0]->{level},
###LogSD					$_[0]->{filename}, $_[0]->{line},
###LogSD					join( "\n\t\t", @print_list ) 	);
###LogSD		use warnings 'uninitialized';
###LogSD	}

###LogSD	1;