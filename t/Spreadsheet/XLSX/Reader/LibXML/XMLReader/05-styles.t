#########1 Test File for Spreadsheet::XLSX::Reader::LibXML::XMLDOM::Styles      8#########9
#!/usr/bin/env perl
my ( $lib, $test_file );
BEGIN{
	$ENV{PERL_TYPE_TINY_XS} = 0;
	my	$start_deeper = 1;
	$lib		= 'lib';
	$test_file	= 't/test_files/xl/';
	for my $next ( <*> ){
		if( ($next eq 't') and -d $next ){
			$start_deeper = 0;
			last;
		}
	}
	if( $start_deeper ){
		$lib		= '../../../../../../' . $lib;
		$test_file	= '../../../../../test_files/xl/';
	}
}
$| = 1;

use	Test::Most tests => 27;
use	Test::Moose;
use IO::File;
use XML::LibXML::Reader;
use	MooseX::ShortCut::BuildInstance v1.8 qw( build_instance );
$MooseX::ShortCut::BuildInstance::re_use_classes = 1;;
use Data::Dumper;
use	lib
		'../../../../../../../Log-Shiras/lib',
		$lib,
	;
#~ use Log::Shiras::Switchboard qw( :debug );
###LogSD	my	$operator = Log::Shiras::Switchboard->get_operator(#
###LogSD						name_space_bounds =>{
###LogSD							UNBLOCK =>{
###LogSD								log_file => 'trace.',
###LogSD							},
###LogSD							Test =>{
###LogSD								XMLReader =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'warn',
###LogSD									},
###LogSD								},
###LogSD								parse_element =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'warn',
###LogSD									},
###LogSD								},
###LogSD								get_format_position =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'trace',
###LogSD									},
###LogSD								},
###LogSD								_get_header_and_value =>{
###LogSD									UNBLOCK =>{
###LogSD										log_file => 'trace',
###LogSD									},
###LogSD								},
#~ ###LogSD								parse_element =>{
#~ ###LogSD									UNBLOCK =>{
#~ ###LogSD										log_file => 'warn',
#~ ###LogSD									},
#~ ###LogSD								},
#~ ###LogSD								_build_date =>{
#~ ###LogSD									UNBLOCK =>{
#~ ###LogSD										log_file => 'warn',
#~ ###LogSD									},
#~ ###LogSD								},
#~ ###LogSD								get_format_position =>{
#~ ###LogSD									UNBLOCK =>{
#~ ###LogSD										log_file => 'trace',
#~ ###LogSD									},
#~ ###LogSD								},
#~ ###LogSD								_get_header_and_position =>{
#~ ###LogSD									UNBLOCK =>{
#~ ###LogSD										log_file => 'trace',
#~ ###LogSD									},
#~ ###LogSD								},
#~ ###LogSD								parse_excel_format_string =>{
#~ ###LogSD									UNBLOCK =>{
#~ ###LogSD										log_file => 'warn',
#~ ###LogSD									},
#~ ###LogSD								},
###LogSD							},
###LogSD						},
###LogSD						reports =>{
###LogSD							log_file =>[ Print::Log->new ],
###LogSD						},
###LogSD					);
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::FmtDefault;
use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles;
use	Spreadsheet::XLSX::Reader::LibXML::Error;
$test_file = ( @ARGV ) ? $ARGV[0] : $test_file;
$test_file .= 'styles.xml';
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'trace', message => [ "Test file is: $test_file" ] );
my  ( 
			$test_instance, $capture, $x, @answer, $error_instance, $format_instance, $file_handle, $coercion,
	);
my 			$row = 0;
my 			@class_attributes = qw(
				file
				error_inst
				format_inst
			);
my  		@class_methods = qw(
				get_format_position
				get_default_format_position
				get_file
				set_file
				has_file
				clear_file
				error
				set_error
				clear_error
				set_warnings
				if_warn
				parse_element
			);
				#~ get_number_format
###LogSD		$phone->talk( level => 'info', message => [ "easy questions ..." ] );
lives_ok{
			$error_instance		= 	Spreadsheet::XLSX::Reader::LibXML::Error->new(
										#~ should_warn => 1,
										should_warn => 0,# to turn off cluck when the error is set
									);
			$format_instance	=  	Spreadsheet::XLSX::Reader::LibXML::FmtDefault->new(
										epoch_year	=> 1904,
										error_inst	=> $error_instance,
				###LogSD				log_space	=> 'Test',
									);
			$test_instance		=	build_instance(
										package => 'TestInstance',
										superclasses => [ 'Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles' ],
										format_inst  => $format_instance,
										file		 => $test_file,
										error_inst	 => $error_instance,
				###LogSD				log_space	=> 'Test',
									);
}										"Prep a new Styles instance";
map{ 
has_attribute_ok
			$test_instance, $_,			"Check that Spreadsheet::XLSX::Reader::LibXML::XMLDOM::Styles has the -$_- attribute"
} 			@class_attributes;
map{
can_ok		$test_instance, $_,
} 			@class_methods;

###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );
ok			$coercion = $format_instance->parse_excel_format_string( '[$-409]d-mmm-yy;@' ),#'(#,##0_);[Red](#,##0)'
										"Create a number conversion from an excel format string";
#~ explain									Dumper( $format_instance );
			my $answer = '12-Sep-05';
is			$coercion->assert_coerce( 37145 ), $answer, #coercecoerce
										"... and see if it returns: $answer";
is			$test_instance->get_format_position( 2, 'cell_coercion' )->{cell_coercion}->display_name, 'Excel_date_164',
										"Check that the excel number coercion at format position 2 is named: Excel_date_164";
is			$test_instance->get_default_format_position->{cell_fill}->{patternFill}->{patternType}, 'none',
										"Check that the default format for fill is: none";
###LogSD		$phone->talk( level => 'debug', message => [ $test_instance->get_format_position( 7, 'cell_font' ) ] );
is			$test_instance->get_format_position( 7, 'cell_font' )->{cell_font}->{sz}, 14,
										"Check that number format position 7 has a font size set to: 14";
is			$test_instance->get_format_position( 7, 'cell_font' )->{cell_font}->{name}, 'Calibri',
										"Check that number format position 7 has a font type set to: Calibri";
lives_ok{
			$test_instance		=	build_instance(
										package => 'TestInstance',
										superclasses => [ 'Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles' ],
										format_inst  => $format_instance,
										file		 => $test_file,
										error_inst	 => $error_instance,
										cache_positions => 0,
				###LogSD				log_space	=> 'Test',
									);
}										"Prep a new Styles instance - without caching";
is			$test_instance->get_format_position( 2, 'cell_coercion' )->{cell_coercion}->display_name, 'Excel_date_164',
										"Check that the excel number coercion at format position 2 is named: Excel_date_164";
###LogSD		$phone->talk( level => 'debug', message => [ $test_instance->get_format_position( 7, 'cell_font' ) ] );
is			$test_instance->get_default_format_position->{cell_fill}->{patternFill}->{patternType}, 'none',
										"Check that the default format for fill is: none";
is			$test_instance->get_format_position( 7, 'cell_font' )->{cell_font}->{sz}, 14,
										"Check that number format position 7 has a font size set to: 14";
is			$test_instance->get_format_position( 7, 'cell_font' )->{cell_font}->{name}, 'Calibri',
										"Check that number format position 7 has a font type set to: Calibri";
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