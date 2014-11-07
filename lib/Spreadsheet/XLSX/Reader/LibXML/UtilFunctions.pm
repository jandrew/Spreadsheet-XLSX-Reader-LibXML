package Spreadsheet::XLSX::Reader::LibXML::UtilFunctions;
use version; our $VERSION = qv('v0.5_1');

use	5.010;
use	Moose::Role;
requires qw(
	get_log_space		
);
use Types::Standard qw( is_Int );
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



=head1 SYNOPSIS
	
	#!perl
	package MyPackage;
	
	###########################
	# SYNOPSIS Screen Output
	# 01: (2, 2)
	###########################
	
=head2 Attributes

Attiributes are ways to change the instances behaviour and can be set as arguments 
to -E<gt>new

=head3 count_from_zero

=over

B<Definition:> A boolean attribute that determines if the numerical output of 
L<parse_column_row|/parse_column_row( $excel_row_id )> provides a response counting from 
Zero or One. True = count from Zero.

B<Accepts:> $bool = (1|0)

=back
	
=head2 Methods

Methods are object methods (not functional methods)

=head3 parse_column_row( $excel_row_id, $count_from_one )

=over

B<Definition:> This is the way to turn an alpha numeric Excel cell ID into row and column 
integers.  If count_from_zero = 1 but you want (column, row) pairs returned counting from 
1 then set $count_from_one = 1.  Or leave it blank to have the pair returned in the format 
defined by L<count_from_zero|/count_from_zero>

B<Accepts:> $excel_row_id, $count_from_one

B<Returns:> ( $column_number, $row_number ) - integers

=back

=head3 build_cell_label( $column, $row, $count_from_one )

=over

B<Definition:> This is the way to turn a (column, row) pair into an excel ID.  If 
$count_from_one is set then the ($column, $row pair will be treated at counting from one 
independant of how L<count_from_zero|/count_from_zero> is set.
integers

B<Accepts:> $column, $row, $count_from_one (in that order and position)

B<Returns:> ( $excel_cell_id ) - integers

=back

=head3 counting_from_zero( $bool )

=over

B<Definition:> This turns on (or off) counting from zero where the alternative is to 
count from 1.

B<Accepts:> $bool = (1|0)

B<Returns:> nothing

=back

=head3 get_excel_position( $int )

=over

B<Definition:> If you wish to use this sheet agnostically of the L<count_from_zero|/count_from_zero> 
setting then you can use this method to translate integers to a count-from-one number.  No action is 
taken if the attribute is set to 0.

B<Accepts:> a $count_from_one or a $count_from_zero int

B<Returns:> a $count_from_one int

=back

=head3 get_used_position( $int )

=over

B<Definition:> If you wish to use this sheet agnostically of the L<count_from_zero|/count_from_zero> 
setting then you can use this method to translate integers from a count-from-one number to whatever 
scheme is in force from the attribute.  No action is taken if the attribute is set to 0.

B<Accepts:> a $count_from_one int

B<Returns:> a $count_from_one or a $count_from_zero int

=back

=head1 SUPPORT

=over

L<github Spreadsheet-XLSX-Reader-LibXML/issues|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

=back

=head1 TODO

=over

B<1.> Add a read raw text between tags step in there somewhere

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

L<version>

L<Moose::Role>

L<Types::Standard>

requires

	name_space
	set_error

=back

=head1 SEE ALSO

=over

L<Spreadsheet::XLSX>

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=back

=cut

#########1#########2 main pod documentation end  5#########6#########7#########8#########9