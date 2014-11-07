package Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData;
use version; our $VERSION = qv('v0.5_1');

use	Moose::Role;
use 5.010;
requires qw(
	node_name	byte_consumed	move_to_first_att	move_to_next_att
	inner_xml	next_element	node_depth			node_value
);
###LogSD	use Log::Shiras::Telephone;

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9



#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9



#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub parse_element{
	my ( $self, $level ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					($self->get_log_space .  '::parse_element' ), );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Parsing element: " . $self->node_name,
	###LogSD			".. at byte position: " . $self->byte_consumed,
	###LogSD			(( defined $level ) ? "..to level: $level" : undef ),] );
	my $current_level //= $self->node_depth;
	my $current_ref;
	
	# Load the attributes	
	my $result = $self->move_to_first_att;
	###LogSD	$phone->talk( level => 'trace', message => [
	###LogSD		"Result of the first attribute move: $result",
	###LogSD		".. at byte position: " . $self->byte_consumed,
	###LogSD		'..for node name: ' . $self->node_name			] );
	ATTRIBUTELIST: while( $result > 0 ){
		my $att_name = $self->node_name;
		my $att_value = $self->node_value;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Reading attribute: $att_name", "..and value: $att_value" ] );
		if( $att_name eq 'val' ){
			$current_ref = $att_value;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Assuming we are at the bottom of the node with a found attribute val: $current_ref"] );
			last ATTRIBUTELIST;
		}else{
			$current_ref->{$att_name} = "$att_value";
		}
		$result = $self->move_to_next_att;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Result of the move: $result",
		###LogSD		".. at byte position: " . $self->byte_consumed, ] );
	}
	my $node_text;
	$node_text = $self->inner_xml;
	if( defined( $node_text ) and length( $node_text ) > 0 and $node_text !~ /^</ ){
		$current_ref->{raw_text} = $node_text;
		delete $current_ref->{'xml:space'};
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Perl ref to this point: ", $current_ref ] );
	
	# Stop or go down another level
	my ( $hash_ref, $list_ref );
	if( defined $level and $level <= $self->node_depth ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Not going down to next sibling because of limit: $level",
		###LogSD		'libxml2 current level: ' . $self->node_depth,
		###LogSD		"..acting at current level: $current_level",
		###LogSD		'..at byte position: ' . $self->byte_consumed, ] );
	}else{
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Pre dive node name: " . $self->node_name ] );
		$result = $self->next_element;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Attempted to go deeper with result: $result",
		###LogSD		(($self->node_name) ? ('Current node name: ' . $self->node_name) : undef),
		###LogSD		'libxml2 current level: ' . $self->node_depth,
		###LogSD		'parser current level: ' . $current_level,
		###LogSD		'..at byte position: ' . $self->byte_consumed,
		###LogSD		((defined $level) ? "..for max allowed level: $level" : undef),] );
		SUBNODES: while( ($self->node_depth - 1) == $current_level ){
			my $node_name = $self->node_name;
			my $byte_count = $self->byte_consumed;
			my $sub_ref = $self->parse_element( ($level) ? $level : undef );
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Returned from parse element with: ", $sub_ref, ] );
			push @$list_ref, $sub_ref;
			$hash_ref->{ $node_name} = $sub_ref;
			###LogSD	$phone->talk( level => 'info', message => [
			###LogSD		"Coallated nodes to this point:", $list_ref, $hash_ref,
			###LogSD		(($self->node_name) ? ('current libxml2 node name: ' . $self->node_name) : undef),
			###LogSD		"current libxml2 node level: " . $self->node_depth,
			###LogSD		"passed level: " . $current_level,
			###LogSD		((defined $level) ? "..against max level: $level" : undef),
			###LogSD		"Bytes consumed: $byte_count"] );
			
			# Go down as possible
			#~ my $not_indexed = 1;
			while( (( $self->node_depth - 1 ) > $current_level) ){
					#~ or ($not_indexed and !ref $sub_ref)				){figure this out!!!
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		'Attempting to find the next node at level: ' . ( $current_level + 1 ), ] );
				$result = $self->next_element;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Node index result: $result",
				###LogSD		(($self->node_name) ? ('libxml2 current node name: ' . $self->node_name) : undef),
				###LogSD		'libxml2 current node depth: ' . $self->node_depth ] );
				#~ $not_indexed = 0;
			}
			
			# Go up when finished
			if( $self->node_depth <= $current_level ){
				###LogSD	$phone->talk( level => 'info', message => [
				###LogSD		'Reached the end of node group at level: ' . ($current_level+1) ] );
				last SUBNODES;
			}
		}
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Finished node walking with attribute ref:", $current_ref,
	###LogSD		'..list ref:', $list_ref,
	###LogSD		'..and hash ref:', $hash_ref,] );
	
	# Determine what to return
	if( $list_ref ){
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"Resolving Node list: ", $list_ref,
		###LogSD		'..ref count: ' .  scalar( @$list_ref ),
		###LogSD		'or Hash ref:', $hash_ref,
		###LogSD		'..hash ref values count: ' . scalar( values( %$hash_ref ) ),] );
		if( $current_ref and ( keys $current_ref )[0] eq 'count' ){
			$current_ref->{list} = $list_ref;
		}elsif( scalar( @$list_ref ) == scalar( values( %$hash_ref ) ) ){
			@$current_ref{ keys %$hash_ref } = ( values( %$hash_ref ) );
		}else{
			$current_ref->{list} = $list_ref;
		}
	}elsif( !$current_ref ){
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"No node list to process",] );
		$current_ref = 1;
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Returning ref: ", $current_ref ] );
	return $current_ref;
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9



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

The default of this module, however, is to count from zero.  Meaning that cell A1 is 
equal to (0, 0).  See the L<Attributes|/Attributes> and L<Methods|/Methods> section for 
ways to change this behaviour.

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