package Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData;
use version; our $VERSION = qv('v0.34.4');

use	Moose::Role;
use 5.010;
requires qw(
	node_name		byte_consumed	move_to_first_att	move_to_next_att
	node_depth		node_value		node_type			has_value	
	start_reading
);#text_value
###LogSD	requires 'get_log_space';
###LogSD	use Log::Shiras::Telephone;

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9



#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9



#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub parse_element{
	my ( $self, $level ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					($self->get_log_space .  '::parse_element' ), );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Parsing element: " . ($self->node_name//''),
	###LogSD			(( defined $level ) ? "..to level: $level" : undef ),] );
	
	my( $success, $return ) = $self->_parse_element( $level );
	return $return;
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

sub _parse_element{
	my ( $self, $level ) = @_;
	my	$current_level = $self->node_depth;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					($self->get_log_space .  '::parse_element' ), );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Parsing element: " . ($self->node_name//''),
	###LogSD			"Node type: " . ($self->node_type//''),
	###LogSD			".. at byte position: " . $self->byte_consumed,
	###LogSD			".. at node depth: $current_level",
	###LogSD			(( defined $level ) ? "..to level: $level" : undef ),] );
	
	# Check for a text node type (and return immediatly if so)
	if( $self->has_value ){
		my $node_text = $self->node_value;
		###LogSD		$phone->talk( level => 'info', message =>[
		###LogSD			"This is a text node - returning value: $node_text",] );
		return ( $self->start_reading, $node_text );
	}
	
	# Load the attributes
	my $current_ref;
	my $result = $self->move_to_first_att;
	###LogSD	$phone->talk( level => 'trace', message => [
	###LogSD		"Result of the first attribute move: $result",
	###LogSD		".. at byte position: " . $self->byte_consumed,
	###LogSD		'..for node name: ' . ($self->node_name//'')	] );
	ATTRIBUTELIST: while( $result > 0 ){
		my $att_name = $self->node_name;
		my $att_value = $self->node_value;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Reading attribute: $att_name", "..and value: $att_value" ] );
		if( $att_name eq 'val' ){
			$current_ref = $att_value;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Assuming we are at the bottom of the attribute list with a found attribute val: $current_ref"] );
			last ATTRIBUTELIST;
		}else{
			$current_ref->{$att_name} = "$att_value";
		}
		$result = $self->move_to_next_att;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Result of the move: $result",
		###LogSD		".. at byte position: " . $self->byte_consumed, ] );
	}
	
	# Advance to the next node to see if we proceed
	my ( $hash_ref, $list_ref, $value_case, $duplicate_keys ) = ( undef, undef, 0, 0 );#, $exceed_depth
		$result		= $self->start_reading;
	my	$node_depth = $self->node_depth;
	my	$node_name  = $self->node_name;
		$node_name	= ($node_name and $node_name eq '#text') ? 'raw_text' : $node_name;
	my $byte_count = $self->byte_consumed;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Initial node advance from level -$current_level- produces result: $result",
	###LogSD		"..and advanced to libxml2 level: $node_depth",
	###LogSD		((defined $level) ? "..for max allowed level: $level" : undef),
	###LogSD		(($self->node_name) ? ('Current node name: ' . $self->node_name) : undef),
	###LogSD		'..at byte position: ' . $self->byte_consumed, '..with type: ' . $self->node_type ] );
	
	# Decide how to proceed
	if( $result ){
		
		# Stop or go down another level
		if( $current_level >= $node_depth ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"The last node was a flat node - returning: $result", $current_ref ] );
			return ( $result, $current_ref );
		}elsif( defined $level and ( $level + 1 ) <= $node_depth ){
			#~ $exceed_depth = 1;
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Not going down to next sibling because of limit: $level",
			###LogSD		"libxml2 current level: $node_depth",
			###LogSD		"..acting at current level: $current_level",
			###LogSD		'..at byte position: ' . $self->byte_consumed, ] );
			return ( $result, $current_ref );
		}else{
			my $counter = 0;
			SUBNODES: while( $result and $node_depth >= ($current_level+1) ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		'Reading subnode: ' . $counter++ ] );
				( $result, my $sub_ref ) =
					$self->_parse_element( ($level) ? $level : undef );
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Returned from parse element with: ", $sub_ref, ] );
				if( defined $sub_ref or $node_name eq 'v' ){
					$value_case = 1;
				}
				$duplicate_keys = 1 if exists $hash_ref->{$node_name};
				push @$list_ref, ( $sub_ref // $node_name );
				$hash_ref->{$node_name} = ($sub_ref // 1 );
				$node_depth = $self->node_depth;
				###LogSD	$phone->talk( level => 'info', message => [
				###LogSD		"Coallated nodes to this point:", $list_ref, $hash_ref,
				###LogSD		(($self->node_name) ? ('current libxml2 node name: ' . $self->node_name) : undef),
				###LogSD		'..libxml2 node type: ' . $self->node_type,
				###LogSD		"current libxml2 node level: $node_depth" ,
				###LogSD		"subnode level: " . ($current_level + 1),
				###LogSD		((defined $level) ? "..against max level: $level" : undef),
				###LogSD		"Bytes consumed: $byte_count", "last result: $result" ] );
				if( $self->node_type == 15 ){
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		'Found an end tag' ] );
					$result		= $self->start_reading;
				}
				
				# come back up to current level as needed
				my $node_type = $self->node_type;
				while( $node_type == 15 or ($result and $node_depth > ($current_level + 1)) ){#!$sub_ref or 
						#~ or ($not_indexed and !ref $sub_ref)				){figure this out!!!
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		'Attempting to find the next node at level: ' . ( $current_level + 1 ), ] );
					$result		= $self->start_reading;
					$node_depth = $self->node_depth;
					$node_type  = $self->node_type;
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Node index result: $result",
					###LogSD		"libxml2 current node depth: $node_depth",
					###LogSD		"And node type: $node_type"					] );
				}
				$node_name	= $self->node_name;
				$node_depth = $self->node_depth;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		(($self->node_name) ? "libxml2 current node name: $node_name" : undef),
				###LogSD		"Node type: " . ($self->node_type//''),
				###LogSD		'libxml2 current node depth: ' . $self->node_depth,
				###LogSD		"Against current node list level: $current_level",	] );
			}
		}
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Finished node walking with attribute ref:", $current_ref,
		###LogSD		'..list ref:', $list_ref,
		###LogSD		'..and hash ref:', $hash_ref,] );
	}else{
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"Reached the end of the file!" ] );
	}
	
	# Determine what to return
	if( $list_ref ){
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"Resolving Node list: ", $list_ref,
		###LogSD		'..ref count: ' .  scalar( @$list_ref ),
		###LogSD		'or Hash ref:', $hash_ref,
		###LogSD		'..hash ref values count: ' . scalar( values( %$hash_ref ) ),
		###LogSD		"With value case: $value_case", "..and duplicate keys: $duplicate_keys" ] );
		
		if( (!exists $current_ref->{count} or $current_ref->{count} != 1) and 
			$value_case and !$duplicate_keys 								){
			@$current_ref{ keys( %$hash_ref ) } = ( values( %$hash_ref ) );
			delete $current_ref->{'xml:space'} if exists $current_ref->{raw_text};
		}else{
			$current_ref->{list} = $list_ref;
		}
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"Current ref resolved to:", $current_ref,] );
	}
	
	# Handle empty Excel shared string values
	if( $current_ref and ref( $current_ref ) and
		(!exists $current_ref->{t} or $current_ref->{t} eq 'str') and
		exists $current_ref->{v} and
		$current_ref->{v} == 1			){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Identified an empty string" ] );
		$current_ref->{v} = {raw_text => ''};
		delete $current_ref->{t};
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Returning ref: ", $current_ref ] );
	return ( $result, $current_ref );
}

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9



#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose::Role;
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData - 
XMLReader to turn xlsx XML to perl hashes
    
=head1 DESCRIPTION

B<This documentation is written to explain ways to extend this package.  To use the data 
extraction of Excel workbooks, worksheets, and cells please review the documentation for  
L<Spreadsheet::XLSX::Reader::LibXML>,
L<Spreadsheet::XLSX::Reader::LibXML::Worksheet>, and 
L<Spreadsheet::XLSX::Reader::LibXML::Cell>>

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
