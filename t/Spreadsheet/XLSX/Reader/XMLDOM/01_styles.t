#########1 Test File for Spreadsheet::XLSX::Reader::XMLDOM::Styles    7#########8#########9
#!perl
$| = 1;

use	Test::Most;
use	Test::Moose;
use	MooseX::ShortCut::BuildInstance v1.8 qw( build_instance );
use	lib
		'../../../../../../Log-Shiras/lib',
		'../../../../../lib',;
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
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::XMLDOM::Styles;
use	Spreadsheet::XLSX::Reader::Error;
my	$test_file = ( @ARGV ) ? $ARGV[0] : '../../../../test_files/xl/';
	$test_file .= 'styles.xml';
###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space => 'main', );
###LogSD		$phone->talk( level => 'trace', message => [ "Test file is: $test_file" ] );
my  ( 
			$test_instance, $capture, $x, @answer, $error_instance,
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
				encoding
				get_number_format
				get_font_definition
			);
###LogSD		$phone->talk( level => 'info', message => [ "easy questions ..." ] );
map{ 
has_attribute_ok
			'Spreadsheet::XLSX::Reader::XMLDOM::Styles', $_,
										"Check that Spreadsheet::XLSX::Reader::XMLDOM::Styles has the -$_- attribute"
} 			@class_attributes;
map{
can_ok		'Spreadsheet::XLSX::Reader::XMLDOM::Styles', $_,
} 			@class_methods;

###LogSD		$phone->talk( level => 'info', message => [ "harder questions ..." ] );
lives_ok{
			$test_instance	=	Spreadsheet::XLSX::Reader::XMLDOM::Styles->new(
									file_name	=> $test_file,
									log_space	=> 'Test',
									error_inst	=> Spreadsheet::XLSX::Reader::Error->new(
										#~ should_warn => 1,
										should_warn => 0,# to turn off cluck when the error is set
									),
									epoch_year	=> 1904,
								);
}										"Prep a new Styles instance";

###LogSD		$phone->talk( level => 'info', message => [ "hardest questions ..." ] );
is			$test_instance->get_number_format( 2 )->{translation}->name, 'C164FromNum',
										"Check that the custom formatted number translation is named: C164FromNum";
is			$test_instance->get_number_format( 7 )->{font}->getChildrenByTagName('sz')->[0]->getAttribute( 'val' ), 14,
										"Check that number format position 7 has a font size set to: 14";
is			$test_instance->get_font_definition( 1 )->getChildrenByTagName('color')->[0]->getAttribute( 'rgb' ), 'FFFF0000',
										"Check that font definition position 1 has the color set to: FFFF0000";
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