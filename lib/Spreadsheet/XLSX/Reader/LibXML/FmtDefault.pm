package Spreadsheet::XLSX::Reader::LibXML::FmtDefault;
use version; our $VERSION = qv('v0.5_1');

use	5.010;
use	Moose::Role;
requires qw(
	get_log_space
);

use Types::Standard qw( InstanceOf ArrayRef Str );
use lib	'../../../../../lib',;
###LogSD	use Log::Shiras::Telephone;

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9



#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has excel_region =>(
		isa		=> Str,
		default	=> 'en',
		reader	=> 'get_excel_region',
	);
	
has	target_encoding =>(
		isa			=> Str,
		reader		=> 'get_target_encoding',
		writer		=> 'set_target_encoding',
		default		=> 'UTF-8',
		required	=> 1,
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub	change_output_encoding{
	my ( $self, $string, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD				name_space 	=> $self->get_log_space . '::change_output_encoding', );
	###LogSD		$phone->talk( level => 'info', message => [
	###LogSD				"Changing the encoding of: $string",
	###LogSD				'..to encoding type: ' . $self->get_target_encoding ] );
	return $string;
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _defined_excel_translations =>(
		isa		=> ArrayRef,
		traits	=> ['Array'],
		default	=> sub{ [
						'General',
						'0',
						'0.00',
						'#,##0',
						'#,##0.00',
						'$#,##0_);($#,##0)',
						'$#,##0_);[Red]($#,##0)',
						'$#,##0.00_);($#,##0.00)',
						'$#,##0.00_);[Red]($#,##0.00)',
						'0%',
						'0.00%',
						'0.00E+00',
						'# ?/?',
						'# ??/??',
						'yyyy-m-d',      # Was 'm-d-yy', which is bad as system default
						'd-mmm-yy',
						'd-mmm',
						'mmm-yy',
						'h:mm AM/PM',
						'h:mm:ss AM/PM',
						'h:mm',
						'h:mm:ss',
						'm-d-yy h:mm',
						undef, undef, undef, undef, undef, undef, undef, undef,
						'#,##0_);(#,##0)',
						'#,##0_);[Red](#,##0)',
						'#,##0.00_);(#,##0.00)',
						'#,##0.00_);[Red](#,##0.00)',
						'_(*#,##0_);_(*(#,##0);_(*"-"_);_(@_)',
						'_($*#,##0_);_($*(#,##0);_($*"-"_);_(@_)',
						'_(*#,##0.00_);_(*(#,##0.00);_(*"-"??_);_(@_)',
						'_($*#,##0.00_);_($*(#,##0.00);_($*"-"??_);_(@_)',
						'mm:ss',
						'[h]:mm:ss',
						'mm:ss.0',
						'##0.0E+0',
						'@'
					]
		},
		reader => 'get_defined_excel_format_list',
		writer => 'set_defined_excel_format_list',
		handles =>{
			get_defined_excel_format => 'get',
			total_defined_excel_formats => 'count',
		},
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9



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