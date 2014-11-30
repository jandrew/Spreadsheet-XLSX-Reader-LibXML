package Spreadsheet::XLSX::Reader::LibXML::UtilFunctions;
use version; our $VERSION = qv('v0.16.2');

use	5.010;
use	Moose::Role;
use	Carp qw( confess );
requires qw(
	get_log_space		
);
use Types::Standard qw( is_Int is_Num );
use lib	'../../../../../../lib',;
###LogSD	use Log::Shiras::Telephone;

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9



#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9



#########1 Public Methods     3#########4#########5#########6#########7#########8#########9



#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9



#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _add_integer_separator{
	my ( $self, $int, $comma, $frequency ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD				name_space 	=> $self->get_log_space . '::_util_function::_add_integer_separator', );
	###LogSD		$phone->talk( level => 'info', message => [
	###LogSD				"Attempting to add the separator -$comma- to " . 
	###LogSD				"the integer portion of: $int" ] );
		$comma //= ',';
	my	@number_segments;
	if( is_Int( $int ) ){
		while( $int =~ /(-?\d+)(\d{$frequency})$/ ){
			$int= $1;
			unshift @number_segments, $2;
		}
		unshift @number_segments, $int;
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		'Final parsed list:', @number_segments ] );
		return join( $comma, @number_segments );
	}else{
		###LogSD	$phone->talk( level => 'warn', message => [
		###LogSD		"-$int- is not an integer!" ] );
		return undef;
	}
}

sub _continuous_fraction{# http://www.perlmonks.org/?node_id=41961
	my ( $self, $decimal, $max_iterations, $max_digits ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD				name_space 	=> $self->get_log_space . '::_util_function::_continuous_fraction', );
	###LogSD		$phone->talk( level => 'info', message => [
	###LogSD				"Attempting to build an integer fraction with decimal: $decimal",
	###LogSD				"Using max iterations: $max_iterations",
	###LogSD				"..and max digits: $max_digits",			] );
	my	@continuous_integer_list;
	my	$start_decimal = $decimal;
	confess "Passed bad decimal: $decimal" if !is_Num( $decimal );
	while( $max_iterations > 0 and ($decimal >= 0.00001) ){
		$decimal = 1/$decimal;
		( my $integer, $decimal ) = $self->_integer_and_decimal( $decimal );
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"The integer of the inverse decimal is: $integer",
		###LogSD		"The remaining decimal is: $decimal" ] );
		if($integer > 999 or ($decimal < 0.00001 and $decimal > 1e-10) ){
			###LogSD	$phone->talk( level => 'info', message => [
			###LogSD		"Either I found a large integer: $integer",
			###LogSD		"...or the decimal is small: $decimal" ] );
			if( $integer <= 999 ){
				push @continuous_integer_list, $integer;
			}
			last;
		}
		push @continuous_integer_list, $integer;
		$max_iterations--;
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"Remaining iterations: $max_iterations" ] );
	}
	###LogSD	$phone->talk( level => 'info', message => [
	###LogSD		"The current continuous fraction integer list is:", @continuous_integer_list ] );
	my ( $numerator, $denominator ) = $self->_integers_to_fraction( @continuous_integer_list );
	if( !$numerator or ( $denominator and length( $denominator ) > $max_digits ) ){
		my $denom = 9 x $max_digits;
		my ( $int, $dec ) = $self->_integer_and_decimal( $start_decimal * $denom );
		$int++;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Passing through the possibilities with start numerator: $int",
		###LogSD		"..and start denominator: $denom", "Against start decimal: $decimal"] );
		my $lowest = ( $start_decimal >= 0.5 ) ?
				{ delta => (1-$start_decimal), numerator => 1, denominator => 1 } :
				{ delta => ($start_decimal-0), numerator => 0, denominator => 1 } ;
		while( $int ){
			my @check_list;
			my $low_int = $int - 1;
			my $low_denom = int( $low_int/$start_decimal ) + 1;
			push @check_list,
					{ delta => abs( $int/$denom - $start_decimal ), numerator => $int, denominator => $denom },
					{ delta => abs( $low_int/$denom - $start_decimal ), numerator => $low_int, denominator => $denom },
					{ delta => abs( $low_int/$low_denom - $start_decimal ), numerator => $low_int, denominator => $low_denom },
					{ delta => abs( $int/$low_denom - $start_decimal ), numerator => $int, denominator => $low_denom };
			my @fixed_list = sort { $a->{delta} <=> $b->{delta} } @check_list;
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		'Built possible list of lower fractions:', @fixed_list ] );
			if( $fixed_list[0]->{delta} < $lowest->{delta} ){
				$lowest = $fixed_list[0];
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		'Updated lowest with:', $lowest ] );
			}
			$int = $low_int;
			$denom = $low_denom - 1;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Attempting new possibilities with start numerator: $int",
			###LogSD		"..and start denominator: $denom", "Against start decimal: $decimal"] );
		}
		($numerator, $denominator) = $self->_best_fraction( @$lowest{qw( numerator denominator )} );
	}
	###LogSD	$phone->talk( level => 'info', message => [
	###LogSD		(($numerator) ? "Final numerator: $numerator" : undef),
	###LogSD		(($denominator) ? "Final denominator: $denominator" : undef), ] );
	if( !$numerator ){
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"Fraction is below the finite value - returning undef" ] );
		return undef;
	}elsif( !$denominator or $denominator == 1 ){
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"Rounding up to: $numerator" ] );
		return( $numerator );
	}else{
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"The final fraction is: $numerator/$denominator" ] );
		return $numerator . '/' . $denominator;
	}
}

# Takes a list of terms in a continued fraction, and converts them
# into a fraction.
sub _integers_to_fraction {# ints_to_frac
	my ( $self, $numerator, $denominator) = (shift, 0, 1); # Seed with 0 (not all elements read here!)
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD				name_space 	=> $self->get_log_space . '::_util_function::_integers_to_fraction', );
	###LogSD		$phone->talk( level => 'info', message => [
	###LogSD				"Attempting to build an integer fraction with the continuous fraction list: " .
	###LogSD				join( ' - ', @_ ), "With a seed numerator of -0- and seed denominator of -1-" ] );
	for my $integer( reverse @_ ){# Get remaining elements
		###LogSD	$phone->talk( level => 'info', message => [ "Now processing: $integer" ] );
		($numerator, $denominator) =
			($denominator, $integer * $denominator + $numerator);
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"New numerator: $numerator", "New denominator: $denominator", ] );
	}
	($numerator, $denominator) = $self->_best_fraction($numerator, $denominator);
	###LogSD	$phone->talk( level => 'info', message => [
	###LogSD		"Updated numerator: $numerator",
	###LogSD		(($denominator) ? "..and denominator: $denominator" : undef) ] );
	return ( $numerator, $denominator );
}


# Takes a numerator and denominator, in scalar context returns
# the best fraction describing them, in list the numerator and
# denominator
sub _best_fraction{#frac_standard 
	my ($self, $n, $m) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD				name_space 	=> $self->get_log_space . '::_util_function::_best_fraction', );
	###LogSD		$phone->talk( level => 'info', message => [
	###LogSD				"Finding the best fraction", "Start numerator: $n", "Start denominator: $m" ] );
	$n = $self->_integer_and_decimal($n);
	$m = $self->_integer_and_decimal($m);
	###LogSD	$phone->talk( level => 'info', message => [ 
	###LogSD		"Updated numerator and denominator ( $n / $m )" ] );
	my $k = $self->_gcd($n, $m);
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
	my ( $self, $decimal ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD				name_space 	=> $self->get_log_space . '::_util_function::_integer_and_decimal', );
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
	my ($self, $n, $m) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD				name_space 	=> $self->get_log_space . '::_util_function::_gcd', );
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

no Moose::Role;
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::UtilFunctions - A useful Role for number mashing
    
=head1 DESCRIPTION

These functions are taken from various sources.  They are not really meant to be used 
outside of the package or by the end user.  They can be especially usefull for number 
conversions though.

POD not complete!

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