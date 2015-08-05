package Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData;
use version; our $VERSION = qv('v0.38.8');
###LogSD	warn "You uncovered internal logging statements for Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData-$VERSION";

use	Moose::Role;
use	Carp 'confess';
use 5.010;
requires qw(
	node_name		byte_consumed	move_to_first_att	move_to_next_att
	node_depth		node_value		node_type			has_value	
	start_reading
);#text_value
use Types::Standard qw( is_StrictNum );
###LogSD	requires 'get_log_space';
###LogSD	use Log::Shiras::Telephone;

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9



#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9



#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub parse_element{
	my ( $self, @args ) = @_;
	confess "Passed too many arguments to 'parse_element': " . join( '~|~', @args ) if @args > 2;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					($self->get_log_space .  '::parse_element' ), );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Parsing element: " . ($self->node_name//''),
	###LogSD			( @args ? ("...with args: ", @args) : undef ),] );
	my	$args;
	for my $arg ( @args ){
		if( is_StrictNum( $arg ) ){
			$args->{level} = $arg;
		}elsif( $arg ){
			$args->{select_tag} = $arg;
		}
	}
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Parsing element: " . ($self->node_name//''),
	###LogSD			( $args ? ("...with args: ", $args) : undef ),] );
	my ( $success, $return );
	if( exists $args->{select_tag} ){
		( $success, $return ) = $self->_get_tag_string( $args->{select_tag} );
	}elsif( exists $args->{level} ){	
		( $success, $return ) = $self->_parse_element( $args->{level} );
	}else{	
		( $success, $return ) = $self->_parse_element;
	}
	return $return;
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9



#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _parse_element{
	my ( $self, $level ) = @_;
	my	$current_level = $self->node_depth;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					($self->get_log_space .  '::parse_element::_parse_element' ), );
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
				$hash_ref->{$node_name} = $sub_ref;
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
					eval '$self->start_reading';
					$result = 1;
					###LogSD	$phone->talk( level => 'info', message => [
					###LogSD		"After read;",
					###LogSD		(($self->node_name) ? ('current libxml2 node name: ' . $self->node_name) : undef),
					###LogSD		'..libxml2 node type: ' . $self->node_type,
					###LogSD		"current libxml2 node level: " . $self->node_depth, ] );
					if( $@ ){
						###LogSD	$phone->talk( level => 'info', message => [
						###LogSD		"Start reading failed with message:", $@ ] );
						$node_depth = 0;
						last SUBNODES;
					}
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
		}elsif( (scalar( keys( %$hash_ref ) ) == 1 and !$duplicate_keys) ){
			$current_ref = $hash_ref;
		}else{
			$current_ref->{list} = $list_ref;
		}
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"Current ref resolved to:", $current_ref,] );
	}
	
	# Handle empty Excel shared string values
	###LogSD	$phone->talk( level => 'info', message => [
	###LogSD		"Current ref resolved to:", $current_ref,] );
	if( $current_ref and ref( $current_ref ) ){
		if( (!exists $current_ref->{t} or !$current_ref->{t} or $current_ref->{t} eq 'str') and
			exists $current_ref->{v} and !$current_ref->{v}				){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Identified an empty string" ] );
			$current_ref->{v} = {raw_text => undef};
			delete $current_ref->{t};
		}elsif( exists $current_ref->{t} and !$current_ref->{t} ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"badly formed space record" ] );
			$current_ref->{t} = {raw_text => $current_ref->{raw_text}};
			delete $current_ref->{raw_text};
			delete $current_ref->{'#text'};
		}
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Returning ref: ", $current_ref ] );
	return ( $result, $current_ref );
}

sub _get_tag_string{
	my ( $self, $target_node) = @_;
	my	$current_level = $self->node_depth;
	my	$node_name = ($self->node_name//'');
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					($self->get_log_space .  '::parse_element::_get_tag_string' ), );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Parsing element: $node_name",
	###LogSD			"Node type: " . ($self->node_type//''),
	###LogSD			"... from node depth: $current_level",
	###LogSD			"... collecting values for: $target_node", ] );
	
	# Check for the case where the target node is the current node
	my	$node_text = '';
	if( $node_name eq $target_node ){
		$node_text = $self->node_value;
		###LogSD		$phone->talk( level => 'info', message =>[
		###LogSD			"This is the target text node - returning value: $node_text",] );
		return ( $self->start_reading, $node_text );
	}
	
	# Iterate through this node
	my	$result		= $self->start_reading;
	my	$node_depth = $self->node_depth;
	my	$node_type  = $self->node_type;
		$node_name	= $self->node_name;
	my	$last_match = 0;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Node index result: $result",
	###LogSD		"libxml2 current node depth: $node_depth",
	###LogSD		"And node type: $node_type",
	###LogSD		"And node name: $node_name",
	###LogSD		"And last match state: $last_match"		] );
	while( $node_depth > $current_level ){
		###LogSD		$phone->talk( level => 'debug', message =>[
		###LogSD			"continuing the journey",] );
		if( $last_match and $self->has_value ){
			$node_text .= $self->node_value;
			###LogSD		$phone->talk( level => 'info', message =>[
			###LogSD			"This is a target text node - total text: $node_text",] );
			$last_match = 0;
		}elsif( $node_name eq $target_node ){
			$last_match = 1;
			###LogSD	$phone->talk( level => 'info', message =>[
			###LogSD		"Found a target node: $node_name",] );
		}else{
			$last_match = 0;# only take the next node
		}
		$result		= $self->start_reading;
		$node_depth = $self->node_depth;
		$node_name	= $self->node_name;
		$node_type  = $self->node_type;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Node index result: $result",
		###LogSD		"libxml2 current node depth: $node_depth",
		###LogSD		"And node type: $node_type"	,
		###LogSD		"And node name: $node_name",
		###LogSD		"And last match state: $last_match",
		###LogSD		"...with current string |$node_text|"		] );
		last if !$result;
	}
	while( $result and $node_type eq '15' ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Advancing past the end tags", ] );
		$result		= $self->start_reading;
		$node_type  = $self->node_type;
		###LogSD	$node_depth = $self->node_depth;
		###LogSD	$node_name	= $self->node_name;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Node index result: $result",
		###LogSD		"libxml2 current node depth: $node_depth",
		###LogSD		"And node type: $node_type"	,
		###LogSD		"And node name: $node_name"		] );
	}
	###LogSD	$phone->talk( level => 'info', message =>[
	###LogSD		"Returning text: $node_text",] );
	return ( $result, $node_text );
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose::Role;
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData - 
XMLReader to turn xlsx XML to perl hashes

=head1 SYNOPSIS

	#!/usr/bin/env perl
	use Data::Dumper;
	use	MooseX::ShortCut::BuildInstance qw( build_instance );
	use	Spreadsheet::XLSX::Reader::LibXML::XMLReader;
	use	Spreadsheet::XLSX::Reader::LibXML::Error;
	use	Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData;
	my  $test_file = 'xl/sharedStrings.xml';
	my  $test_instance = build_instance(
			package => 'TestIntance',
			superclasses =>[ 'Spreadsheet::XLSX::Reader::LibXML::XMLReader', ],
			add_roles_in_sequence =>[ 'Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData', ],
			file	=> $test_file,
			error_inst	=> Spreadsheet::XLSX::Reader::LibXML::Error->new,
		);
	map{ $test_instance->next_element( 'si' ) }( 0..15 );# Go somewhere interesting
	print Dumper( $test_instance->parse_element ) . "\n";

	###############################################
	# SYNOPSIS Screen Output
	# 01: $VAR1 = {
	# 02:           'list' => [
	# 03:                       {
	# 04:                         't' => {
	# 05:                                'raw_text' => 'He'
	# 06:                              }
	# 07:                       },
	# 08:                       {
	# 09:                         'rPr' => {
	# 10:                                  'color' => {
	# 11:                                             'rgb' => 'FFFF0000'
	# 12:                                           },
	# 13:                                  'sz' => '11',
	# 14:                                  'b' => 1,
	# 15:                                  'scheme' => 'minor',
	# 16:                                  'rFont' => 'Calibri',
	# 17:                                  'family' => '2'
	# 18:                                },
	# 19:                         't' => {
	# 20:                                'raw_text' => 'llo '
	# 21:                              }
	# 22:                       },
	# 23:                       {
	# 24:                         'rPr' => {
	# 25:                                  'color' => {
	# 26:                                             'rgb' => 'FF0070C0'
	# 27:                                           },
	# 28:                                  'sz' => '20',
	# 29:                                  'b' => 1,
	# 30:                                  'scheme' => 'minor',
	# 31:                                  'rFont' => 'Calibri',
	# 32:                                  'family' => '2'
	# 33:                                },
	# 34:                         't' => {
	# 35:                                'raw_text' => 'World'
	# 36:                              }
	# 37:                       }
	# 38:                     ]
	# 39:         };
	###############################################
    
=head1 DESCRIPTION

This documentation is written to explain ways to use this module when writing your own excel 
parser.  To use the general package for excel parsing out of the box please review the 
documentation for L<Workbooks|Spreadsheet::XLSX::Reader::LibXML>,
L<Worksheets|Spreadsheet::XLSX::Reader::LibXML::Worksheet>, and 
L<Cells|Spreadsheet::XLSX::Reader::LibXML::Cell>

This package is used convert xml to L<deep|/parse_element( $level, $target_node )> perl data 
structures.  As a note deep perl xml and  data structures are not one for one compatible to xml.  
However, there is a subset of xml that reasonably translates to deep perl structures.  For this 
implementation node names are treated as hash keys unless there are multiple subnodes within a node 
that have the same name.  In this case the subnode name is stripped and each node is added as a 
subref in an arrary ref.  The overall arrayref is attached to the key list.  Attributes are also 
treated as hash keys at the same level as the sub nodes.  Text nodes (or raw text between tags) is 
treated as having the key 'raw_text'.

This reader assumes that it is a role added to a class built on 
L<Spreadsheet::XLSX::Reader::LibXML::XMLReader> it expects to get the methods provided by that type 
of file reader to use to traverse the node.  As a consequence it doesn't accept an xml object since 
it expects the overall file to be read serially.

=head2 Required Methods

L<node_name|Spreadsheet::XLSX::Reader::LibXML::XMLReader/node_name>

L<byte_consumed|Spreadsheet::XLSX::Reader::LibXML::XMLReader/byte_consumed>

L<move_to_first_att|Spreadsheet::XLSX::Reader::LibXML::XMLReader/move_to_first_att>

L<move_to_next_att|Spreadsheet::XLSX::Reader::LibXML::XMLReader/move_next_att>

L<node_depth|Spreadsheet::XLSX::Reader::LibXML::XMLReader/node_depth>

L<node_value|Spreadsheet::XLSX::Reader::LibXML::XMLReader/node_value>

L<node_type|Spreadsheet::XLSX::Reader::LibXML::XMLReader/node_type>

L<has_value|Spreadsheet::XLSX::Reader::LibXML::XMLReader/has_value>

L<start_reading|Spreadsheet::XLSX::Reader::LibXML::XMLReader/start_reading>

=head2 Methods

These are the methods provided by this module.

=head3 parse_element( $level, $target_node )

=over

B<Definition:> This returns a perl equivalent of the xml data structure where the 
L<Spreadsheet::XLSX::Reader::LibXML::XMLReader> is currently positioned.  If there 
is a $target_node provided then it just returns all of the values stored in the full 
node concatenated together.  It is assumed that the desired information are the text 
elements of each node.  ($level is not applied in this case).  If $level is provided 
in the absence of a $target_node then the conversion from xml to deep perl data is only 
carried to that (absolute) level.  Where neither datum is provided a full translation 
to a perl data structure is done.  When this method is done the xml reader will be left 
at the next xml tag after the current full node (even or up).  L<It will clear all end 
tags.>

B<Accepts:> $level ( a positive integer ), $target_node ( the specific target node name 
case sensitive ) - These values are order independant since one is a number  and the 
other a string

B<Returns:> ($success, $data_ref ) This method returns a list with the first element 
being success or failure and the second element being the data ref (or string) corresponding to the 
xml being parsed by L<Spreadsheet::XLSX::Reader::LibXML::XMLReader>.

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

This software is copyrighted (c) 2014, 2015 by Jed Lund

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
