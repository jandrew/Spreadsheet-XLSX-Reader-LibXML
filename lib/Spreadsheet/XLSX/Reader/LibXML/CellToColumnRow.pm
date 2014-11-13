package Spreadsheet::XLSX::Reader::LibXML::CellToColumnRow;
use version; our $VERSION = qv('v0.10.2');

use	Moose::Role;
requires qw(
	get_log_space
	set_error
);
use Types::Standard qw( Bool );
###LogSD	use Log::Shiras::Telephone;

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9

my	$lookup_ref ={
		A => 1, B => 2, C => 3, D => 4, E => 5, F => 6, G => 7, H => 8, I => 9, J => 10,
		K => 11, L => 12, M => 13, N => 14, O => 15, P => 16, Q => 17, R => 18, S => 19,
		T => 20, U => 21, V => 22, W => 23, X => 24, Y => 25, Z => 26,
	};
my	$lookup_list =[ qw( A B C D E F G H I J K L M N O P Q R S T U V W X Y Z ) ];

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9



#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub parse_column_row{
	my ( $self, $cell ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					($self->get_log_space .  '::parse_column_row' ), );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Parsing file row number and file column number from: $cell" ] );
	my ( $column, $row ) = $self->_parse_column_row( $cell );
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"File Column: $column", "File Row: $row" ] );
	###LogSD	use warnings 'uninitialized';
	return( $column, $row );
}

sub build_cell_label{
	my ( $self, $column, $row ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					($self->get_log_space .  '::build_cell_label' ), );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Converting file column -$column- and file row -$row- to a cell ID" ] );
	my $cell_label = $self->_build_cell_label( $column, $row );
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"Cell label is: $cell_label" ] );
	return $cell_label;
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9



#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _parse_column_row{
	my ( $self, $cell ) = @_;
	my ( $column, $error_list_ref );
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					($self->get_log_space .  '::_parse_column_row' ), );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Parsing excel row and column number from: $cell" ] );
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

sub _build_cell_label{
	my ( $self, $column, $row ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					($self->get_log_space .  '::build_cell_label' ), );
	###LogSD	no	warnings 'uninitialized';
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Converting column -$column- and row -$row- to a cell ID" ] );
	###LogSD	use	warnings 'uninitialized';
	my $error_list;
	if( !defined $column ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"The column is not defined" ] );
		$column = '';
		push @$error_list, 'missing column';
	}else{
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Excel column: $column" ] );
		$column -= 1;
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"From zero: $column" ] );
		if( $column > 16383 ){
			push @$error_list, 'column too large';
			$column = '';
		}elsif( $column < 0 ){
			push @$error_list, 'column too small';
			$column = '';
		}else{
			my $first_letter = int( $column / (26 * 26) );
			$column = $column - $first_letter * (26 * 26);
			$first_letter = ( $first_letter ) ? $lookup_list->[$first_letter - 1] : '';
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"First letter is: $first_letter", "New column is: $column" ] );
			my $second_letter = int( $column / 26 );
			$column = $column - $second_letter * 26;
			$second_letter =
				( $second_letter ) ? $lookup_list->[$second_letter - 1] : 
				( $first_letter ne '' ) ? 'A' : '' ;
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Second letter is: $second_letter", "New column is: $column" ] );
			my $third_letter = $lookup_list->[$column];
			$column = $first_letter . $second_letter . $third_letter;
		}
	}
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"Column letters are: $column" ] );
	
	if( !defined $row ){
		$row = '';
		push @$error_list, 'missing row';
	}else{
		if( $row > 1048576 ){
			push @$error_list, 'row too large';
			$row = '';
		}elsif( $row < 1 ){
			push @$error_list, 'row too small';
			$row = '';
		}
	}
	$self->set_error(
		"Failures in build_cell_label include: " . join( ' - ', @$error_list )
	) if $error_list;
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"Row is: $row" ] );
	
	my $cell_label = "$column$row";
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"Cell label is: $cell_label" ] );
	return $cell_label;
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose::Role;
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::CellToRowColumn - Translate Excel cell IDs to row column
    
=head1 DESCRIPTION

This is a fairly simple implementation of a regex and math to find the column and row
position in excel from an 'A1' style Excel cell ID.  It is important to note that column 
letters do not equal digits in a modern 26 position numeral system since the excel 
implementation is effectivly zeroless.

The default of this module is to count from 1 (the excel convention).  Meaning that cell 
A1 is equal to (1, 1).  However, there is a layer of abstraction in order to support 
count from zero settings using the Moose around function.  See the L<Methods|/Methods> 
section for more details on the implementation.

=head1 SYNOPSIS
	
	#!perl
	package MyPackage;
	use Moose;
	with 'Spreadsheet::XLSX::Reader::LibXML::CellToColumnRow';

	sub set_error{}
	sub get_log_space{}
		
	sub my_method{
		my ( $self, $cell ) = @_;
		my ($column, $row ) = $self->parse_column_row( $cell );
		print $self->error if( !defined $column or !defined $row );
		return ($column, $row );
	}

	package main;

	my $parser = MyPackage->new( count_from_zero => 0 );
	print '(' . join( ', ', $parser->my_method( 'B2' ) ) . ")'\n";
	
	###########################
	# SYNOPSIS Screen Output
	# 01: (2, 2)
	###########################
	
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

L<version>

L<Moose::Role>

L<Types::Standard>

=back

=head2 Requires

=over

B<get_log_space>

B<set_error>

=back

=head1 SEE ALSO

=over

L<Spreadsheet::XLSX>

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=back

=cut

#########1#########2 main pod documentation end  5#########6#########7#########8#########9