package Spreadsheet::XLSX::Reader::LibXML::Types;
use version; our $VERSION = qv('v0.10.4');

use strict;
use warnings;
use Type::Utils -all;
use Type::Library 0.046
	-base,
	-declare => qw(
		FileName					XMLFile						XLSXFile
		ParserType					StrictInt					StrictTwoDecimal
		StrictCommaInt				StrictCommaTwoDecimal		StrictPercent
		StrictTwoDecimalPercent		ScientificNotation			NumWithFraction
		ShortUSDate					MediumDate					DayMonth
		MonthYear					TwelveHourMinute			EpochYear
		Alignment					PassThroughType				CellType
		CellID						ReadSpan					PositiveNum
		NegativeNum					ZeroOrUndef					NotNegativeNum
		
		Excel_number_0				OneFromNum					TwoFromNum
		ThreeFromNum				FourFromNum					NineFromNum
		TenFromNum					ElevenFromNum				TwelveFromNum
		FourteenFromWinExcelNum		FourteenFromAppleExcelNum	FifteenFromWinExcelNum
		FifteenFromAppleExcelNum	SixteenFromWinExcelNum		SixteenFromAppleExcelNum
		SeventeenFromWinExcelNum	SeventeenFromAppleExcelNum	EighteenFromNum
	);
BEGIN{ extends "Types::Standard" };
my $try_xs =
		exists($ENV{PERL_TYPE_TINY_XS}) ? !!$ENV{PERL_TYPE_TINY_XS} :
		exists($ENV{PERL_ONLY})         ?  !$ENV{PERL_ONLY} :
		1;
if( $try_xs and exists $INC{'Type/Tiny/XS.pm'} ){
	eval "use Type::Tiny::XS 0.010";
	if( $@ ){
		die "You have loaded Type::Tiny::XS but versions prior to 0.010 will cause this module to fail";
	}
}
use DateTimeX::Format::Excel;
use lib	'../../../../lib',;
###LogSD	use Log::Shiras::Telephone;

#########1 Package Variables  3#########4#########5#########6#########7#########8#########9

my	$win_excel_converter	= DateTimeX::Format::Excel->new;
my	$apple_excel_converter	= DateTimeX::Format::Excel->new( system_type => 'apple_excel' );
our	$log_space				= 'Spreadsheet::XLSX::Reader::Types';

#########1 Type Library       3#########4#########5#########6#########7#########8#########9

declare FileName,
	as Str,
    where{ -r $_ },
    message{ 
        ( $_ ) ? 
            "Could not find / read the file: $_" : 
            'No value passed to the file_name test' 
    };
	
declare XMLFile,
	as Str,
	where{ $_ =~ /\.xml$/ and -r $_  },
	message{
        ( $_ !~ /\.xml$/ ) ?
            "The string -$_- does not have an xml file extension" :
		( !-r $_ ) ?
			"Could not find / read the file: $_" : 
            'No value passed to the xml_file test' ;
    };
	
declare XLSXFile,
	as Str,
	where{ $_ =~ /\.xlsx$/ and -r $_ },
	message{
        ( $_ !~ /\.xlsx$/ ) ?
            "The string -$_- does not have an xlsx file extension" :
		( !-r $_ ) ?
			"Could not find / read the file: $_" : 
            'No value passed to the xlsx_file test' 
    };

my	$parser_definitions = qr/^(dom|reader|sax)$/;
declare ParserType, 
	as Str,
	where{ $_ =~ $parser_definitions },
	message{
        ( $_ ) ? 
            "The string -$_- does not match $parser_definitions" : 
            'No value passed to the ParserType test' 
    };

coerce ParserType,
	from Str,
	via{ lc( $_ ) };
	
declare StrictInt,
	as StrictNum,
	where{ defined( $_ ) and (int( $_ ) == $_) };
	
declare StrictTwoDecimal,
	as StrictNum,
	where{ defined( $_ ) and length( $_ - int( $_ ) ) == 2 };
	
declare StrictCommaInt,
	as Str,
	where{ defined( $_ ) and $_ =~ /^-?[\d,]$/ };
	
declare StrictCommaTwoDecimal,
	as Str,
	where{ defined( $_ ) and length( $_ - int( $_ ) ) == 2 };

declare StrictPercent,
	as Str,
	where{ $_ =~ /^-?\d+%$/ };

declare StrictTwoDecimalPercent,
	as Str,
	where{ $_ =~ /^-?\d+\.\d{2}%$/ };

declare ScientificNotation,
	as Str,
	where{ $_ =~ /^\d\.\d{2} E(\+|\-)\d+$/ };

declare NumWithFraction,
	as Str,
	where{ $_ =~ /^(\d+\s)?\d+\/\d+$/ };
	
declare ShortUSDate,
	as Str,
	where{
		$_ =~ /^(\d{1,2})\/\d{1,2}\/\d{2}$/ and
		$1 > 0 and $1 < 13
	};#(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)
	
declare MediumDate,
	as Str,
	where{
		$_ =~ /^(\d{1,2})\/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\/\d{2}$/ and
		$1 > 0 and $1 < 32
	};#
	
declare DayMonth,
	as Str,
	where{
		$_ =~ /^(\d{1,2})\/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)$/ and
		$1 > 0 and $1 < 32
	};#
	
declare MonthYear,
	as Str,
	where{
		$_ =~ /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\/\d{2}$/
	};#

declare TwelveHourMinute,
	as Str,
	where{ $_ =~ /^\d{1,2}:\d{2} (AM|PM)$/ };

declare EpochYear,
	as Int,
	where{ $_ == 1900 or $_ == 1904 };

declare Alignment,
	as Dict[
		horizontal	=> Maybe[ Str ],
		vertical	=> Maybe[ Str ],
	];

declare PassThroughType,
	as Any;

declare CellType,
	as Enum[ 's', 'number' ];

declare CellID,
	as StrMatch[ qr/^[A-Z]{1,3}\d+$/ ];

declare ReadSpan,
	as Enum[ qw( sheet_span row_span non_null ) ];
	
declare PositiveNum,
	as Num,
	where{ $_ > 0 };

declare NegativeNum,
	as Num,
	where{ $_ < 0 };
	
declare ZeroOrUndef,
	as Maybe[Num],
	where{ !$_ };
	
declare NotNegativeNum,
	as Num,
	where{ $_ > -1 };


#########1 Excel Defined Converions     4#########5#########6#########7#########8#########9

declare_coercion Excel_number_0,
	to_type Any, from Maybe[Any],
	via{ $_ };

declare_coercion OneFromNum,
	to_type StrictInt, from Maybe[Num],
	via{ 
		my $num = $_;
		return '0' if( !defined $num );
		$num =~ /(-?)(\d*)\.?(\d)/;
		my ( $sign, $integer, $first_dec ) = ( $1, $2, $3 );
		###LogSD	my	$phone = Log::Shiras::Telephone->new(
		###LogSD				name_space 	=> $log_space . '::OneFromNum', );
		###LogSD	no warnings 'uninitialized';
		###LogSD		$phone->talk( level => 'info', message => [
		###LogSD				"Coercing num: $num",
		###LogSD				"Into integer using: |$sign| |$integer| |$first_dec|" ] );
		###LogSD	use warnings 'uninitialized';
		$integer += 1 if( $first_dec and $first_dec > 4);
		my $return = "$sign$integer";
		###LogSD		$phone->talk( level => 'info', message => [ "Returning: $return" ] );
		return $return;
	};

declare_coercion TwoFromNum,
	to_type StrictTwoDecimal, from Maybe[Num],
	via{ ( !defined $_ ) ? '0.00' : sprintf( '%01.2f', $_ ) };

declare_coercion ThreeFromNum,
	to_type StrictCommaInt, from Maybe[Num],
	via{ 
		my $num = $_;
		return '0' if( !defined $num );
		$num =~ /(-?)(\d*)\.?(\d)/;
		my ( $return, $integer, $first_dec ) = ( $1, $2, $3 );
		###LogSD	my	$phone = Log::Shiras::Telephone->new(
		###LogSD				name_space 	=> $log_space . '::TwoFromNum', );
		###LogSD	no warnings 'uninitialized';
		###LogSD		$phone->talk( level => 'info', message => [
		###LogSD				"Coercing num: $num",
		###LogSD				"Into threes separated integer using: |$return| |$integer| |$first_dec|" ] );
		###LogSD	use warnings 'uninitialized';
		$integer += 1 if( $first_dec and $first_dec > 4);
		### <where> - integer: $integer
		$integer = _add_threes_separator( $integer );
		$return .= $integer;
		###LogSD		$phone->talk( level => 'info', message => [ "Returning: $return" ] );
		return $return;
	};

declare_coercion FourFromNum,
	to_type StrictCommaTwoDecimal, from Maybe[Num],
	via{
		my $num = $_;
		my $return;
		if( !defined $num ){
			$return = '0.00';
		}else{
			$num =~ /(-?)(\d*)\.?(\d*)/;
			$return = $1;
			my ( $integer, $decimal ) = ( $2, $3 );
			$return .= _add_threes_separator( $integer );
			sprintf( '%.2f', ( "\." . $decimal ) ) =~ /(\.\d+)/;
			$return .= $1;
		}
		return $return
	};

declare_coercion NineFromNum,
	to_type StrictPercent, from Maybe[Num],
	via{ 
		my $num = $_;
		return '0%' if( !defined $num );
		$num = $num*100;
		$num =~ /(-?)(\d*)\.?(\d)/;
		my ( $return, $integer, $first_dec ) = ( $1, $2, $3 );
		$integer += 1 if( $first_dec and $first_dec > 4);
		$return .= $integer . '%';
		return $return;
	};

declare_coercion TenFromNum,
	to_type StrictTwoDecimalPercent, from Maybe[Num],
	via{ 
		my $num = $_;
		return '0.00%' if( !defined $num );
		$num = $num*100;
		$num =~ /(-?)(\d*)\.?(\d*)/;
		my ( $return, $integer, $decimal ) = ( $1, $2, $3 );
		sprintf( '%.2f', ( "\." . $decimal ) ) =~ /(\.\d+)/;
		$return .= $integer . $1 . '%';
		return $return;
	};

declare_coercion ElevenFromNum,
	to_type ScientificNotation, from Maybe[Num],
	via{ 
		my $num = $_;
		return '0.00 E+00' if( !defined $num );
		my	$return .= sprintf '%2.2E', $num;
		return $return;
	};

declare_coercion TwelveFromNum,
	to_type NumWithFraction, from Maybe[Num],
	via{ 
		my $num = $_;
		return '0/0' if( !defined $num );
		my	$absolute	= abs( $num );
		my	$return		= ( $absolute != $num ) ? '-' : '';
			$num		= $absolute;
		my	$integer	= int( $num );
		my	$decimal	= $num - $integer;
		$return .= "$integer" if $integer;
		return $return if !defined $decimal;
		$return .= ' ' if $integer;
		my	$fraction = _continuous_fraction( $decimal, 20 );
		$return .= $fraction;
		return $return;
	};

declare_coercion FourteenFromWinExcelNum,
	to_type ShortUSDate, from Maybe[Num],
	via{
		my $num = $_;
		return undef if( !defined $num );
		$win_excel_converter->parse_datetime( $num )->format_cldr( 'M-d-yy' )
	};

declare_coercion FourteenFromAppleExcelNum,
	to_type ShortUSDate, from Maybe[Num],
	via{
		my $num = $_;
		return undef if( !defined $num );
		$apple_excel_converter->parse_datetime( $num)->format_cldr( 'M-d-yy' )
	};

declare_coercion FifteenFromWinExcelNum,
	to_type MediumDate, from Maybe[Num],
	via{
		my $num = $_;
		return undef if( !defined $num );
		$win_excel_converter->parse_datetime( $num)->format_cldr( 'd-MMM-yy' )
	};

declare_coercion FifteenFromAppleExcelNum,
	to_type MediumDate, from Maybe[Num],
	via{
		my $num = $_;
		return undef if( !defined $num );
		$apple_excel_converter->parse_datetime( $_ )->format_cldr( 'd-MMM-yy' )
	};

declare_coercion SixteenFromWinExcelNum,
	to_type DayMonth, from Maybe[Num],
	via{
		my $num = $_;
		return undef if( !defined $num );
		$win_excel_converter->parse_datetime( $_ )->format_cldr( 'd-MMM' )
	};

declare_coercion SixteenFromAppleExcelNum,
	to_type DayMonth, from Maybe[Num],
	via{
		my $num = $_;
		return undef if( !defined $num );
		$apple_excel_converter->parse_datetime( $_ )->format_cldr( 'd-MMM' )
	};

declare_coercion SeventeenFromWinExcelNum,
	to_type MonthYear, from Maybe[Num],
	via{
		my $num = $_;
		return undef if( !defined $num );
		$win_excel_converter->parse_datetime( $_ )->format_cldr( 'MMM-yy' )
	};

declare_coercion SeventeenFromAppleExcelNum,
	to_type MonthYear, from Maybe[Num],
	via{
		my $num = $_;
		return undef if( !defined $num );
		$apple_excel_converter->parse_datetime( $_ )->format_cldr( 'MMM-yy' )
	};

declare_coercion EighteenFromNum,
	to_type TwelveHourMinute, from Maybe[Num],
	via{
		my $num = $_;
		return undef if( !defined $num );
		$apple_excel_converter->parse_datetime( $_ )->format_cldr( 'h:mm a' )
	};

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9



#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _add_threes_separator{
	my ( $int, $comma ) = @_;
		$comma //= ',';
	my	@number_segments;
	while( $int =~ /(-?\d+)(\d{3})$/ ){
		$int= $1;
		unshift @number_segments, $2;
	}
	unshift @number_segments, $int;
	return join( $comma, @number_segments );
}

sub _continuous_fraction{# http://www.perlmonks.org/?node_id=41961
	my ( $decimal, $max_iterations ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD				name_space 	=> $log_space . '::_continuous_fraction', );
	###LogSD		$phone->talk( level => 'info', message => [
	###LogSD				"Attempting to build an integer fraction with decimal: $decimal",
	###LogSD				"Using max iterations: $max_iterations" ] );
	my	@continuous_integer_list = ();
	my	$run_once = 0;
	while( $max_iterations > 0 and (!$run_once or $decimal > 0.001) ){
		$decimal = 1/$decimal;
		( my $integer, $decimal ) = _integer_and_decimal( $decimal );
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"The integer of the inverse decimal is: $integer",
		###LogSD		"The remaining decimal is: $decimal" ] );
		if( $run_once and ($integer > 999 or ($decimal < 0.00001 and $decimal > 1e-10) )){
			###LogSD	$phone->talk( level => 'info', message => [
			###LogSD		"Either I found a large integer: $integer",
			###LogSD		"...or the decimal is small: $decimal" ] );
			last;
		}else{
			$run_once = 1;
		}
		push @continuous_integer_list, $integer;
		$max_iterations--;
			###LogSD	$phone->talk( level => 'info', message => [
			###LogSD		"Remaining iterations: $max_iterations" ] );
	}
	###LogSD	$phone->talk( level => 'info', message => [
	###LogSD		"The current continuous fraction integer list is:", @continuous_integer_list ] );
	my $fraction = _integers_to_fraction( @continuous_integer_list );
	###LogSD	$phone->talk( level => 'info', message => [
	###LogSD		"The final fraction is: $fraction" ] );
	return $fraction;
}

# Takes a list of terms in a continued fraction, and converts them
# into a fraction.
sub _integers_to_fraction {# ints_to_frac
	my ($numerator, $denominator) = (0, 1); # Seed with 0
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD				name_space 	=> $log_space . '::_integers_to_fraction', );
	###LogSD		$phone->talk( level => 'info', message => [
	###LogSD				"Attempting to build an integer fraction with the continuous fraction list: " .
	###LogSD				join( ' - ', @_ ), "With a seed numerator of -0- and seed denominator of -1-" ] );
	for my $integer( reverse @_ ){
		###LogSD	$phone->talk( level => 'info', message => [ "Now processing: $integer" ] );
		($numerator, $denominator) =
			($denominator, $integer * $denominator + $numerator);
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"New numerator: $numerator", "New denominator: $denominator" ] );
	}
	return _best_fraction($numerator, $denominator);
}


# Takes a numerator and denominator, in scalar context returns
# the best fraction describing them, in list the numerator and
# denominator
sub _best_fraction{#frac_standard 
	my ($n, $m) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD				name_space 	=> $log_space . '::_best_fraction', );
	###LogSD		$phone->talk( level => 'info', message => [
	###LogSD				"Finding the best fraction", "Start numerator: $n", "Start denominator: $m" ] );
	$n = _integer_and_decimal($n);
	$m = _integer_and_decimal($m);
	###LogSD	$phone->talk( level => 'info', message => [ 
	###LogSD		"Updated numerator and denominator ( $n / $m )" ] );
	my $k = _gcd($n, $m);
	###LogSD	$phone->talk( level => 'info', message => [ "Greatest common divisor: $k" ] );
	$n = $n/$k;
	$m = $m/$k;
	###LogSD	$phone->talk( level => 'info', message => [ 
	###LogSD		"Reduced numerator and denominator ( $n / $m )" ] );
	if ($m < 0) {
		###LogSD	$phone->talk( level => 'info', message => [ "the divisor is less than zero" ] );
		$n *= -1;
		$m *= -1;
	}
	$m = undef if $m == 1;
	###LogSD	no warnings 'uninitialized';
	###LogSD	$phone->talk( level => 'info', message => [ 
	###LogSD		"Final numerator and denominator ( $n / $m )" ] );
	###LogSD	use warnings 'uninitialized';
	if (wantarray) {
		return ($n, $m);
	}else {
		return ( $m ) ? "$n/$m" : $n;
	}
}

# Takes a number, returns the best integer approximation and
#	(in list context) the error.
sub _integer_and_decimal {
	my ( $decimal ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD				name_space 	=> $log_space . '::_integer_and_decimal', );
	###LogSD		$phone->talk( level => 'info', message => [ 
	###LogSD				"Splitting integer from decimal for: $decimal" ] );
	my $integer = int( $decimal );
	###LogSD		$phone->talk( level => 'info', message => [ "Integer: $integer" ] );
	if(wantarray){
		return($integer, $decimal - $integer);
	}else{
		return $integer;
	}
}

# Euclidean algorithm for calculating a GCD.
# Takes two integers, returns the greatest common divisor.
sub _gcd {
	my ($n, $m) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD				name_space 	=> $log_space . '::_gcd', );
	###LogSD		$phone->talk( level => 'info', message => [ 
	###LogSD				"Finding the greatest common divisor for ( $n and $m )" ] );
	while ($m) {
		my $k = $n % $m;
		###LogSD	$phone->talk( level => 'info', message => [ 
		###LogSD		"Remainder after division: $k" ] );
		($n, $m) = ($m, $k);
		###LogSD	$phone->talk( level => 'info', message => [ 
		###LogSD		"Updated factors ( $n and $m )" ] );
	}
	return $n;
} 
	

#########1 Phinish            3#########4#########5#########6#########7#########8#########9
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::Types - A type library for the LibXML xlsx reader
    
=head1 DESCRIPTION

POD not written yet!

=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader::LibXML/issues
|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

=back

=head1 TODO

=over

B<1.> Nothing L<yet|/SUPPORT>

=back

=head1 AUTHOR

=over

=item Jed Lund

=item jandrew@cpan.org

=back

=head1 COPYRIGHT

This program is free software; you can redistribute
it and/or modify it under the same terms as Perl itself.

The full text of the license can be found in the
LICENSE file included with this module.

This software is copyrighted (c) 2014 by Jed Lund

=head1 DEPENDENCIES

=over

L<Spreadsheet::XLSX::Reader::LibXML>

=back

=head1 SEE ALSO

=over

L<Spreadsheet::ParseExcel> - Excel 2003 and earlier

L<Spreadsheet::XLSX> - 2007+

L<Spreadsheet::ParseXLSX> - 2007+

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=over

All lines in this package that use Log::Shiras are commented out

=back

=back

=cut

#########1#########2 main pod documentation end  5#########6#########7#########8#########9