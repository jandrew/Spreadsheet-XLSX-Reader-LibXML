package Spreadsheet::XLSX::Reader::CellToColumnRow;
use version; our $VERSION = version->declare("v0.1_1");

use	Moose::Role;
requires qw(
	get_log_space
	set_error
);
use lib	'../../../../lib';
###LogSD	use Log::Shiras::Telephone;# Fix with CPAN release of Log::Shiras

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9

my	$lookup_ref ={
		A => 1, B => 2, C => 3, D => 4, E => 5, F => 6, G => 7, H => 8, I => 9, J => 10,
		K => 11, L => 12, M => 13, N => 14, O => 15, P => 16, Q => 17, R => 18, S => 19,
		T => 20, U => 21, V => 22, W => 23, X => 24, Y => 25, Z => 26,
	};

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9



#########1 Public Methods     3#########4#########5#########6#########7#########8#########9



#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9



#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub parse_column_row{
	my ( $self, $cell ) = @_;
	my ( $column, $error_list_ref );
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					($self->get_log_space .  '::_get_column_row' ), );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Parsing row number and column number from: $cell" ] );
	my	$regex = qr/^([A-Z])?([A-Z])?([A-Z])?([0-9]*)$/;
	my ( $one_column, $two_column, $three_column, $row ) = $cell =~ $regex;
	no	warnings 'uninitialized';
	my	$column_text = $one_column . $two_column . $three_column;
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"Regex result is: ( $one_column, $two_column, $three_column, $row )" ] );
	
	if( !defined $one_column ){
		push @$error_list_ref, "Could not parse the column component from -$cell-";
	}elsif( !defined $two_column ){
		$column = $lookup_ref->{$one_column};
	}elsif( !defined $three_column ){
		$column = $lookup_ref->{$two_column} + 26 * $lookup_ref->{$one_column};
	}else{
		$column = $lookup_ref->{$three_column} + 26 * $lookup_ref->{$two_column} + 26 * 26 * $lookup_ref->{$one_column};
	}
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"Result of initial parse is column text: $column_text",
	###LogSD		"Column number: $column", "Row number: $row" ] );
	if( $column_text and $column > 16384 ){
		push @$error_list_ref, "The column text -$column_text- points to a position at " .
									"-$column- past the excel limit of: 16,384";
		$column = undef;
	}
	if( !defined $row or $row eq '' ){
		push @$error_list_ref, "Could not parse the row component from -$cell-";
		$row = undef;
	}elsif( $row < 1 ){
		push @$error_list_ref, "The requested row cannot be less than one - you requested: $row";
		$row = undef;
	}elsif( $row > 1048576 ){
		push @$error_list_ref, "The requested row cannot be greater than 1,048,576 " .
									"- you requested: $row";
		$row = undef;
	}
	if( $error_list_ref ){
		if( scalar( @$error_list_ref ) > 1 ){
			$self->set_error( "The regex $regex could not match -$cell-" );
		}else{
			$self->set_error( $error_list_ref->[0] );
		}
	}
	###LogSD	no warnings 'uninitialized';
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"Column: $column", "Row: $row" ] );
	use warnings 'uninitialized';
	return( $column, $row );
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose::Role;
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::CellToRowColumn - Translate Excel cell IDs to row column
    
=head1 DESCRIPTION

This is a fairly simple implementation of a regex and math to find the column and row
position in excel from an 'A1' style Excel cell ID.  It is important to note that column 
letters do not equal digits in a modern 26 position numeral system since the excel 
implementation is effectivly zeroless.

=head1 SYNOPSIS
	
	#!perl
	package MyPackage;
	use Moose;
	use lib '../lib';
	with 'Spreadsheet::XLSX::Reader::CellToColumnRow';

	sub set_error{}
	sub get_log_space{}
		
	sub my_method{
		my ( $self, $cell ) = @_;
		my ($column, $row ) = $self->parse_column_row( $cell );
		print $self->error if( !defined $column or !defined $row );
		return ($column, $row );
	}

	package main;

	my $parser = MyPackage->new;
	print '(' . join( ', ', $parser->my_method( 'B2' ) ) . ")'\n";
	
	###########################
	# SYNOPSIS Screen Output
	# 01: (2, 2)
	###########################
	
=head2 Methods

Methods are object methods (not functional methods)

=head3 parse_column_row( $excel_row_id )

=over

B<Definition:> This is the way to turn an alpha numeric Excel cell ID into row and column 
integers

B<Accepts:> $excel_row_id

B<Returns:> ( $column_number, $row_number ) - integers

=back

=head1 SUPPORT

=over

L<github Spreadsheet-XLSX-Reader/issues|https://github.com/jandrew/Spreadsheet-XLSX-Reader/issues>

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

L<version>

L<Moose::Role>

requires

	name_space
	_set_error

=back

=head1 SEE ALSO

=over

L<Spreadsheet::XLSX>

L<Spreadsheet::XLSX::Reader::TempFilter>

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=back

=cut

#########1#########2 main pod documentation end  5#########6#########7#########8#########9