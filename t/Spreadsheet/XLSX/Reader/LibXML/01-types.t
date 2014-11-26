#########1 Test File for Spreadsheet::XLSX::Reader::LibXML::Types     7#########8#########9
#!env perl
my ( $lib, $test_file );
BEGIN{
	$ENV{PERL_TYPE_TINY_XS} = 0;
	my	$start_deeper = 1;
	$lib		= 'lib';
	$test_file	= 't/test_files/';
	for my $next ( <*> ){
		if( ($next eq 't') and -d $next ){
			$start_deeper = 0;
			last;
		}
	}
	if( $start_deeper ){
		$lib		= '../../../../../' . $lib;
		$test_file	= '../../../../test_files/';
	}
}
$| = 1;

use	Test::Most tests => 39;
use	Test::TypeTiny;
#~ use	Test::Moose;
use Data::Dumper;
use Capture::Tiny qw( capture_stderr );
use	lib 
		'../../../../../../Log-Shiras/lib',
		$lib,
	;
#~ use Log::Shiras::Switchboard v0.21 qw( :debug );#
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
use Spreadsheet::XLSX::Reader::LibXML::Types v0.1 qw(
		PassThroughType				FileName					XLSXFile
		XMLFile						ParserType					EpochYear
		Excel_number_0				CellID						PositiveNum
		NegativeNum					ZeroOrUndef					NotNegativeNum
	);
my	@types_list = (
		PassThroughType,			FileName,					XLSXFile,
		XMLFile	,					ParserType,					EpochYear,
		CellID,						PositiveNum,				NegativeNum,
		ZeroOrUndef,				NotNegativeNum,				Excel_number_0,				
	);
my	$test_dir	= ( @ARGV ) ? $ARGV[0] : $test_file;
my	$xlsx_file	= $test_dir . 'TestBook.xlsx';
my	$xml_file	= $test_dir . '[Content_Types].xml';
my  ( 
			$position, $counter, $exception,
	);
my 			$row = 0;
my			$question_ref =[
				[ 1, 2, 'Help', 0x234, undef ],# PassThroughType
				[ $xlsx_file, $xml_file, 'badfile.not', ],# FileName	
				[ $xlsx_file, $xml_file,],# XLSXFile
				[ $xml_file, $xlsx_file,],#~ XMLFile
				[ 'dom', 'reader' , 'badfile.not' ],#~ ParserType
				[ 1900, 1904, 2000 ],#~ EpochYear
				[ 'A1', 'CCC10000', 'A0' ],#~ CellID
				[ 1, 2, 0.1234, -3 ], #~ PositiveNum
				[ -1, -2, -0.1234, 0 ],#~ NegativeNum
				[ 0, undef, 's', 2 ],#~ ZeroOrUndef
				[ 1, 2, 0.1234, 0, -1],#~ NotNegativeNum
				#~ Excel_number_0
			];
my			$answer_ref = [
				[],
				[undef, undef, 'Could not find / read the file: badfile.not', ],
				[undef, 'The string -badfile.not- does not have an xlsx file extension', ],
				[undef, 'The string -badfile.not- does not have an xml file extension', ],
				[undef, undef, 'The string -badfile.not- does not match ', ],
				[undef, undef, '2000 is not an excel epoch', ],
				[undef, undef, '0 is not a cell ID', ],
				[undef, undef, undef, '\-3 is not a positive number', ],
				[undef, undef, undef, '0 is not a negative number', ],
				[undef, undef, 's is not zero (or undef)', '2 is not zero (or undef)', ],
				[undef, undef, undef, undef, '-1 is not a negative number', ],
			];
###LogSD my $phone = Log::Shiras::Telephone->new;
###LogSD	$phone->talk( level => 'debug', message =>[ 'Start your engines ...' ] );
			#~ no strict 'refs';
			my $x = 0;
			for my $x ( 0..($#types_list - 1) ){
			my $type = $types_list[$x];
###LogSD	$phone->talk( level => 'debug', message =>[ "Testing type: $type" ] );
			for my $y ( 0..$#{$question_ref->[$x]} ){
###LogSD	if( $x == 0 and $y==0 ){
###LogSD		$operator->add_name_space_bounds( {
###LogSD			UNBLOCK =>{
###LogSD				log_file => 'trace',
###LogSD			},
###LogSD		} );
###LogSD	}
###LogSD	$phone->talk( level => 'debug', message =>[
###LogSD		'Testing value: ' . (($question_ref->[$x]->[$y]) ? $question_ref->[$x]->[$y] : '') ] );
			if( $answer_ref->[$x]->[$y] ){
should_fail $question_ref->[$x]->[$y], $type;
			}else{
should_pass	$question_ref->[$x]->[$y], $type;
			}
			}
			}
ok			Excel_number_0->assert_coerce( 'jabberwoky' ),
							"A run on the Excel_number_0 coercion";
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