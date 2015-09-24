package Spreadsheet::XLSX::Reader::LibXML::XMLToPerlData;
use version; our $VERSION = qv('v0.38.18');

use	Moose::Role;
use Data::Dumper;
use 5.010;
requires qw(
	get_empty_return_type		get_text_node				get_attribute_hash_ref
	advance_element_position	location_status
);#text_value
use Types::Standard qw(	Int	ArrayRef	is_HashRef 	is_Int );
use Clone qw( clone );
###LogSD	requires 'get_all_space';
###LogSD	use Log::Shiras::Telephone;

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9



#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9



#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub parse_element{
	my ( $self, $level ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::parse_element', );
	###LogSD		$phone->talk( level => 'debug', message =>[ "Parsing current element",] );
	$self->_clear_partial_ref;
	$self->_clear_bread_crumbs;
	my ( $node_depth, $node_name, $node_type ) = $self->location_status;
	if( defined $level ){
		###LogSD	$phone->talk( level => 'debug', message =>[ "Setting max level with (+1)", $level, $node_depth] );
		$self->_set_max_level( $level + $node_depth + 1 );
	}else{
		$self->_clear_max_level;
	}
	
	# Set the seed data
	my	$base_depth	= $node_depth;
	my	$last_level = $node_depth - 1;
	my	$has_value	= 0;
	my	$time		= 'first';
	
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"Start node name: $node_name",
	###LogSD		"..of type: $node_type",
	###LogSD		"..at libxml2 level: $node_depth",
	###LogSD		"..from base depth: $base_depth",
	###LogSD		"..for time: $time",
	###LogSD		(($self->_has_max_level) ? 
	###LogSD			('..with max allowed level: ' . $self->_get_max_level) : ''),] );
	
	my $sub_ref;
	PARSETHELAYERS: while( ($time eq 'first') or ($node_depth > $base_depth) ){
		$time = 'not_first';
		
		# Check for a rewind
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Checking for rewind state with current node depth: $node_depth",
		###LogSD		"..........................against last node depth: $last_level", ] );
		if( $node_depth < $last_level ){
			my $rewind = $last_level - $node_depth;
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		"Just moved back up: $rewind", ] );
			$self->_rewind( $rewind );
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		"Rewound to depth: $node_depth", ] );
			$last_level = $node_depth;
		}
		
		# Record progress
		if( $node_depth == $last_level ){# Stack the same level
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		"Found another node at level: $node_depth", ] );
			$self->_stack( $node_name );
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		"Finished stacking with :", $self->_get_partial_ref, $self->_get_bread_crumbs ] );
		}else{
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		"Building out: $node_name", ] );
			$self->_accrete_node( $node_name );
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		"Finished accreting with :", $self->_get_partial_ref, ] );
		}
		$has_value = 0;# Reset value tracker for node level
		
		$last_level = $node_depth;
		if( !$self->_has_max_level or $self->_get_max_level > $node_depth ){# Check for the bottom
			# Check for a text node
			my( $result, $node_text, ) = $self->get_text_node;
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		"Finished text node search with:", $result, $node_text, ] );
			
			# If no text node check for an attribute_ref
			my $next_ref;
			if( $result ){
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"That was a text node - adding |$node_text|" ] );
				$has_value = 1;
				$self->_accrete_node( $node_text, 'text' );
				$last_level++;
				###LogSD	$phone->talk( level => 'debug', message =>[ "Loaded the text node"] );
			}else{
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"Not a text node - Checking for an attribute ref" ] );
				( $result, $next_ref ) = $self->get_attribute_hash_ref;
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"The attribute hash_ref search result -$result- produced the ref:", $next_ref ] );
				if( $result ){
					my	$ref_type = ( ref $next_ref eq 'HASH' ) ? undef : 'text';
					###LogSD	$phone->talk( level => 'trace', message => [
					###LogSD		"Adding attribute hash ref type:", $ref_type, $next_ref,	] );
					$has_value = 1;
					$self->_accrete_node( $next_ref, $ref_type );
					$last_level++;
				}
			}
		}
		
		my $move = 'first';
		while(	$move eq 'first' or 
				($self->_has_max_level and $self->_get_max_level < $node_depth) ){
			$move = 'not_first';
			# Move one step forward
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		"Advancing to the next node",	] );
			my $result  = $self->advance_element_position;
			last PARSETHELAYERS if !$result;
			( $node_depth, $node_name, $node_type ) = $self->location_status;
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Result of the next node process: $result", 
			###LogSD		"Next node name: $node_name",
			###LogSD		"..of type: $node_type",
			###LogSD		"..at libxml2 level: $node_depth",
			###LogSD		"..against base depth: $base_depth",
			###LogSD		(($self->_has_max_level) ? 
			###LogSD			('..and for max allowed level: ' . $self->_get_max_level) : ''),] );
		}
		
		# Handle a self contained node with no sub data
		if( $last_level > $node_depth and $has_value == 0 ){
			###LogSD	$phone->talk( level => 'info', message =>[
			###LogSD		"Found a self contained node with no sub-data",	] );
			$last_level += $self->_empty_node( $node_name );
			$has_value = 1;
		}
	}
	
	# Handle a self contained node with no sub data one last time
	if( $has_value == 0 and $last_level > $node_depth ){
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"One final self contained node with no sub-data",	] );
		$self->_empty_node( $node_name );
	}
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"Finished the loop with:", $self->_get_partial_ref, ] );
	
	# Rewind one last time
	my $rewind = $self->_ref_depth - 2;
	###LogSD	$phone->talk( level => 'trace', message => [
	###LogSD		"Rewinding finally: $rewind", ] );
	my $eof_flag = $self->_rewind( $rewind );
	if( $eof_flag ){
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD		"Found -EOF-", ] );
		return $eof_flag;
	}else{
		###LogSD	$phone->talk( level => 'trace', message =>[
		###LogSD		"Finished processing with:", $self->_get_partial_ref ] );
		return( $self->_last_ref_level );#$self->_last_bread_crumb, 
	}
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has	_max_level =>(
		isa			=> 	Int,
		reader		=> '_get_max_level',
		writer		=> '_set_max_level',
		predicate	=> '_has_max_level',
		clearer		=> '_clear_max_level',
	);

has	_partial_ref =>(
		isa			=> 	ArrayRef,
		traits		=> ['Array'],
		reader		=> '_get_partial_ref',
		writer		=> '_set_partial_ref',
		clearer		=> '_clear_partial_ref',
		handles =>{
			_add_ref_level 	=> 'push',
			_set_ref_level 	=> 'set',
			_get_ref_level 	=> 'get',
			_ref_depth		=> 'count',
			_last_ref_level => 'pop',
		}
	);

has	_bread_crumbs =>(
		isa			=> 	ArrayRef,
		traits		=> ['Array'],
		reader		=> '_get_bread_crumbs',
		writer		=> '_set_bread_crumbs',
		clearer		=> '_clear_bread_crumbs',
		handles =>{
			_add_bread_crumb	=> 'push',
			_set_bread_crumb 	=> 'set',
			_get_bread_crumb 	=> 'get',
			_crumb_trail_length	=> 'count',
			_last_bread_crumb	=> 'pop',
		}
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _accrete_node{
	my ( $self, $node_id, $node_type ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::parse_element::_accrete_node', );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Accreting the node:", $node_id, $node_type, $self->_get_partial_ref] );
	my $ref =
		( $node_type and $node_type eq 'text' ) ? $node_id :
		( ref $node_id ) ? $node_id : { $node_id => undef };
	$self->_add_ref_level( $ref );
	$self->_add_bread_crumb( $node_id );
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"Current node representations:", $self->_get_partial_ref, $self->_get_bread_crumbs,] );
}

sub _rewind{
	my ( $self, $rewind_count ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::parse_element::_rewind', );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Rewinding: $rewind_count", $self->_get_bread_crumbs->[-1], $self->_get_partial_ref->[-1], ] );
	if( !defined $rewind_count or $rewind_count < 0 ){
		$self->set_error( "Can't rewind |$rewind_count| times! - possible end of file" );
		return 'EOF'
	}elsif( $rewind_count == 0 ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Skipping rewind since -0- rewinds called!" ] );
	}else{
		$self->_last_bread_crumb;
		my	$current_value	= $self->_last_ref_level;
		my ( $next_value, $current_ref );
		for my $count ( 1..$rewind_count ){
			$next_value		= $self->_last_bread_crumb;
			$current_ref	= $self->_last_ref_level;
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Rewinding:", $current_value, "from current ref:", $current_ref, ] );
			if( ref $current_ref eq 'HASH' ){
				if( exists $current_ref->{list} ){
					if( $current_value =~ /^\s+/ and is_HashRef( $current_ref->{list}->[$next_value - 1] ) ){
						###LogSD	$phone->talk( level => 'debug', message =>[
						###LogSD		"The value:", $current_value, "..is a string of spaces and and doesn't belong in a list of hashrefs",] );
						pop @{$current_ref->{list}};
					}else{
						###LogSD	$phone->talk( level => 'debug', message =>[
						###LogSD		"Pushing:", $current_value, "..at position: $next_value", "on the list in:", $current_ref, ] );
						$current_ref->{list}->[$next_value] = $current_value;
					}
				}else{
					if( $next_value eq 'raw_text' and scalar( keys %$current_ref ) > 1){# Check for bad spaces before a tag
						###LogSD	$phone->talk( level => 'debug', message =>[
						###LogSD		"Found a bad space after a tag with value |$current_value|" ] );
						delete $current_ref->{$next_value};
						$next_value = (keys %$current_ref)[0];# Choose the first random one in the absence of a better decision
					}else{
						###LogSD	$phone->talk( level => 'debug', message =>[
						###LogSD		"Loading:", $current_value, "to key -$next_value- in:", $current_ref, ] );
						$current_ref->{$next_value} = $current_value;
					}
				}
			}else{
				confess "I don't know how to rewind -" . ref $current_ref . "- types";
			}
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Updated current ref:", $current_ref ] );
			$current_value = $current_ref;
		}
		
		if( $current_ref ){
			#Reload last pop
			$self->_add_bread_crumb( $next_value );
			$self->_add_ref_level( $current_ref );
		}
	}
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"Get final node representations:", $self->_get_partial_ref, $self->_get_bread_crumbs,] );
	return undef;
}

sub _stack{
	my ( $self, $node_id ) = @_;
	my	$replace_key	= $self->_last_bread_crumb;
	my	$current_value	= $self->_last_ref_level;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::parse_element::_stack', );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Stacking node id: " . ($node_id//''),( ( is_Int( $replace_key )  ?  "..to position: " :  "..to key: ") . $replace_key) , "in:",$current_value ] );
	my ( $alt_key, $alt_value );
	
	if( $replace_key eq 'raw_text' ){# Check for bad spaces before a tag
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Found a bad space before the tag -$node_id- with value |$current_value->{$replace_key}|" ] );
		delete $current_value->{$replace_key};
		$current_value->{$node_id} = undef;
		$replace_key = $node_id;
	}else{
		if( ref $current_value eq 'HASH' and exists $current_value->{$node_id}){# Change from a hash to a list
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Changing the current level hash to have a list reference from: $replace_key", ] );
			$current_value->{list}->[0] = ($current_value->{$node_id} // $node_id);
			delete $current_value->{$node_id};
		}
		
		if( exists $current_value->{list} ){
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Pushing -$node_id- to the array" ] );
			push @{$current_value->{list}}, $node_id;
			$replace_key = $#{$current_value->{list}};
		}elsif( exists $current_value->{count} ){
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Starting a list to support count of: " . $current_value->{count}, ] );
			$current_value->{list}->[0] = $node_id;
			$replace_key = 0;
		}else{
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"managing a hash ref:", $current_value, $replace_key ] );
			if( !ref $replace_key ){
				$current_value->{$replace_key} = $current_value->{$replace_key};
			}
			$current_value->{$node_id} = undef;
			$replace_key = $node_id;
			if( $node_id eq 'raw_text' and exists $current_value->{'xml:space'} ){
				delete $current_value->{'xml:space'};
			}
		}
	}
	#Reload last pop
	$self->_add_bread_crumb( $replace_key );
	$self->_add_ref_level( $current_value );
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"Get final node representations:", $self->_get_partial_ref, $self->_get_bread_crumbs,] );
	return ( $replace_key, $current_value );
}

sub _empty_node{
	my ( $self, $node_name ) = @_;# Need to do a '_has_class_space' to get class space transition
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::parse_element::_empty_node', );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"Handling an empty node: $node_name", ] );
	my $final_value =
			( $self->get_empty_return_type and $self->get_empty_return_type eq 'empty_string' ) ? '' : undef;
	my $ref = { $node_name => $final_value };
	$self->_add_ref_level( $ref );
	$self->_add_bread_crumb( $ref );
	###LogSD	$phone->talk( level => 'info', message =>[
	###LogSD		"Updated master stack:", $self->_get_partial_ref, ] );
	return 1;
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose::Role;
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::XMLToPerlData - 
Role to turn xlsx XML to perl hashes

=head1 SYNOPSIS

	#!/usr/bin/env perl
	use Data::Dumper;
	use	MooseX::ShortCut::BuildInstance qw( build_instance );
	use	Spreadsheet::XLSX::Reader::LibXML::XMLReader;
	use	Spreadsheet::XLSX::Reader::LibXML::Error;
	use	Spreadsheet::XLSX::Reader::LibXML::XMLToPerlData;
	my  $test_file = '../../../../test_files/xl/sharedStrings.xml';
	my  $test_instance	=	build_instance(
			package => 'TestIntance',
			superclasses =>[ 'Spreadsheet::XLSX::Reader::LibXML::XMLReader', ],
			add_roles_in_sequence =>[ 'Spreadsheet::XLSX::Reader::LibXML::XMLToPerlData', ],
			file => $test_file,
			error_inst => Spreadsheet::XLSX::Reader::LibXML::Error->new,
			add_attributes =>{
				empty_return_type =>{
					reader => 'get_empty_return_type',
				},
			},
		);
	$test_instance->advance_element_position( 'si', 16 );# Go somewhere interesting
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
    
=head1 DESCRIPTION  ############## Re-write XMLReader POD too!!!

This documentation is written to explain ways to use this module when writing your own excel 
parser.  To use the general package for excel parsing out of the box please review the 
documentation for L<Workbooks|Spreadsheet::XLSX::Reader::LibXML>,
L<Worksheets|Spreadsheet::XLSX::Reader::LibXML::Worksheet>, and 
L<Cells|Spreadsheet::XLSX::Reader::LibXML::Cell>

This package is used convert xml to deep perl data structures.  As a note deep perl xml and  
data structures are not one for one compatible to xml.  However, there is a subset of xml that 
reasonably translates to deep perl structures.  For this implementation node names are treated 
as hash keys unless there are multiple subnodes within a node that have the same name.  In this 
case the subnode name is stripped and each node is added as a subref in an arrary ref.  The overall 
arrayref is attached to the key list.  Attributes are also treated as hash keys at the same level 
as the sub nodes.  Text nodes (or raw text between base tags) is treated as having the key 'raw_text'.

This reader assumes that it is a role that can be added to a class built on 
L<Spreadsheet::XLSX::Reader::LibXML::XMLReader> it expects to get the methods provided by that type 
of xml reader for traversing and reading the file.  As a consequence it doesn't accept an xml object 
or file since it expects to access the method below.

=head2 Required Methods

Follow the links to see details of the current implementation.

=over
	
L<get_empty_return_type|Spreadsheet::XLSX::Reader::LibXML::XMLReader/get_empty_return_type>

L<get_text_node|Spreadsheet::XLSX::Reader::LibXML::XMLReader/get_text_node>

L<get_attribute_hash_ref|Spreadsheet::XLSX::Reader::LibXML::XMLReader/get_attribute_hash_ref>

L<advance_element_position|Spreadsheet::XLSX::Reader::LibXML::XMLReader/advance_element_position>

L<location_status|Spreadsheet::XLSX::Reader::LibXML::XMLReader/location_status>

=back

=head2 Methods

These are the methods provided by this module.

=head3 parse_element( $level )

=over

B<Definition:> This returns a deep perl data structure that represents the full xml 
down as many levels as indicated by $level (positive is deeper) or  to the bottom for 
no passed value.  When this method is done the xml reader will be left at the begining 
of the next xml node after the ending flag for the requested node.

B<Accepts:> $level ( a positive integer )

B<Returns:> ($success, $data_ref ) This method returns a list with the first element 
being success or failure and the second element being the data ref corresponding to the 
xml being parsed.

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
