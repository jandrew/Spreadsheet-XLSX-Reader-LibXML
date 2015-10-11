package Spreadsheet::XLSX::Reader::LibXML::XMLReader;
use version; our $VERSION = version->declare('v0.38.22');

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Types::Standard qw(
		Int				HasMethods			Bool
		Num				Str
    );
use XML::LibXML::Reader;
use Data::Dumper;
use Carp 'confess';
use lib	'../../../../../lib',;
###LogSD	with 'Log::Shiras::LogSpace';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use Spreadsheet::XLSX::Reader::LibXML::Types qw(
		IOFileType
	);
use	Spreadsheet::XLSX::Reader::LibXML::Error;

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has file =>(
		isa			=> IOFileType,
		reader		=> 'get_file',
		writer		=> 'set_file',
		predicate	=> 'has_file',
		clearer		=> 'clear_file',
		coerce		=> 1,
		required	=> 1,
		trigger		=> \&_build_xml_reader,
		handles 	=> [ 'close' ],
	);

has	error_inst =>(
		isa			=> 	HasMethods[qw(
							error set_error clear_error set_warnings if_warn
						) ],
		clearer		=> '_clear_error_inst',
		reader		=> '_get_error_inst',
		required	=> 1,
		handles =>[ qw(
			error set_error clear_error set_warnings if_warn
		) ],
		default => sub{ Spreadsheet::XLSX::Reader::LibXML::Error->new },
	);

has	xml_version =>(
		isa			=> 	Num,
		reader		=> 'version',
		writer		=> '_set_xml_version',
		clearer		=> '_clear_xml_version',
	);

has	xml_encoding =>(
		isa			=> 	Str,
		reader		=> 'encoding',
		predicate	=> 'has_encoding',
		writer		=> '_set_xml_encoding',
		clearer		=> '_clear_xml_encoding',
	);

has	xml_header =>(
		isa			=> 	Str,
		reader		=> 'get_header',
		writer		=> '_set_xml_header',
	);

has position_index =>(
		isa			=> Int,
		reader		=> 'where_am_i',
		writer		=> 'i_am_here',
		clearer		=> 'clear_location',
		predicate	=> 'has_position',
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9


sub start_the_file_over{
	my( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::XMLReader::FromFile::start_the_file_over', );
	if( $self->has_file ){
		###LogSD		$phone->talk( level => 'debug', message =>[ "Resetting the XML file" ] );
		#~ $self->_go_to_the_end;
		$self->_close_the_sheet;
		#~ $self->_clear_xml_parser;
		$self->clear_location;
		my $fh = $self->get_file;
		$fh->seek( 0, 0 );
		$self->_set_xml_parser( XML::LibXML::Reader->new( IO => $fh ) );
		return 1;
	}else{
		###LogSD		$phone->talk( level => 'info', message =>[ "No file to reset" ] );
		return undef;
	}
}

sub get_text_node{
	my ( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::XMLReader::get_text_node', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"getting the text value of the node", ] );
	
	# Check for a text node type (and return immediatly if so)
	if( $self->_has_value ){
		my $node_text = $self->_node_value;
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"This is a text node - returning value: $node_text",] );
		return ( 1, $node_text, );
	}
	# Return undef for no value
	return ( undef,);
}

sub get_attribute_hash_ref{
	my ( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::XMLReader::get_attribute_hash_ref', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Extract all attributes as a hash ref", ] );
	
	my $attribute_ref = {};
	my $result = $self->_move_to_first_att;
	###LogSD	$phone->talk( level => 'trace', message =>[
	###LogSD		"Result of the first attribute move: $result",] );
	ATTRIBUTELIST: while( $result > 0 ){
		my $att_name = $self->_node_name;
		my $att_value = $self->_node_value;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Reading attribute: $att_name", "..and value: $att_value" ] );
		if( $att_name eq 'val' ){
			$attribute_ref = $att_value;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Assuming we are at the bottom of the attribute list with a found attribute val: $att_value"] );
			last ATTRIBUTELIST;
		}else{
			$attribute_ref->{$att_name} = "$att_value";
		}
		$result = $self->_move_to_next_att;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Result of the move: $result", ] );
	}
	$result = ( ref $attribute_ref ) ? (keys %$attribute_ref) : 1;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Returning attribute ref:", $attribute_ref ] );
	return ( $result, $attribute_ref );
}

sub advance_element_position{
	my ( $self, $element, $position ) = @_;
	if( $position and $position < 1 ){
		confess "You can only advance element position in a positive direction, |$position| is not correct.";
	}
	$position ||= 1;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::XMLReader::advance_element_position', );
	###LogSD	$phone->talk( level => 'info', message => [
	###LogSD		"Advancing to element -" . ($element//'') . "- -$position- times", ] );
	my ( $result, $node_depth, $node_name, $node_type, $byte_count );
	my $x = 0;
	for my $y ( 1 .. $position ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Advancing position case: $y", ] );
		if( defined $element ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Searching for element: $element", ] );
			$result = $self->_next_element( $element );
		}else{
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Indexing one more generic node", ] );
			( $result, my( $node_depth, $node_name, $node_type ) ) = $self->_next_node;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Received the result: $result", "..at depth: $node_depth",
			###LogSD		"..and node named: $node_name", "..of node type: $node_type" ] );
			
			# Climb out of end tags
			while( $result and $node_type == 15 ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Advancing from end node", ] );
				( $result, $node_depth, $node_name, $node_type ) = $self->_next_node;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Received the result: $result", "..at depth: $node_depth",
				###LogSD		"..and node named: $node_name", "..of node type: $node_type" ] );
			}
		}
		last if !$result;
		$x++;
	}
	if( defined $node_type and $node_type == 0 ){
		###LogSD	$phone->talk( level => 'info', message =>[ "Reached the end of the file!" ] );
	}elsif( !$result ){
		###LogSD	$phone->talk( level => 'info', message =>[
		###LogSD		"Unable to location position -$position- for element: " . ($element//'') ] );
	}else{
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"Actually advanced -$x- positions with result: $result",
		###LogSD		"..indicated by:", $self->location_status ] );
	}
	return $result;
}

sub location_status{
	my ( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::XMLReader::location_status', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Getting the status for the current position", ] );
	my ( $node_depth, $node_name, $node_type ) = ( $self->_node_depth, $self->_node_name, $self->_node_type );
	$node_name	= 
		( $node_type == 0 ) ? 'EOF' :
		( $node_name eq '#text') ? 'raw_text' :
		$node_name;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Currently at libxml2 level: $node_depth",
	###LogSD		"Current node name: $node_name",
	###LogSD		"..for type: $node_type" ] );
	return ( $node_depth, $node_name, $node_type );
}
	

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _xml_reader =>(
	isa			=> 'XML::LibXML::Reader',
	reader		=> '_get_xml_parser',
	writer		=> '_set_xml_parser',
	predicate	=> '_has_xml_parser',
	clearer		=> '_clear_xml_parser',
	handles	=>{
		copy_current_node	=> 'copyCurrentNode',
		_close_the_sheet	=> 'close',
		_node_depth			=> 'depth',
		_node_type			=> 'nodeType',
		_node_name			=> 'name',
		_encoding			=> 'encoding',
		_version			=> 'xmlVersion',
		_next_element		=> 'nextElement',
		_node_value			=> 'value',
		_has_value			=> 'hasValue',
		_move_to_first_att	=> 'moveToFirstAttribute',
		_move_to_next_att	=> 'moveToNextAttribute',
		_read_next_node		=> 'read',
		#~ _go_to_the_end		=> 'finish',
		get_node_all		=> 'readOuterXml',
	},
	trigger => \&_reader_init,
);

has _read_unique_bits =>(
		isa		=> Bool,
		reader	=> '_get_unique_bits',
		writer	=> '_need_unique_bits',
		clearer	=> '_clear_read_unique_bits',
		default	=> 1,
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _next_node{
	my( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::XMLReader::_next_node', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Reading the next node in the xml document", ] );
	my $result = eval{ $self->_read_next_node } ? 1 : 0 ;# Handle unclosed xml tags without dying
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Result of the read: $result", ] );
	my ( $node_depth, $node_name, $node_type ) = $self->location_status;
	if( $node_name eq '#document' and $node_depth == 0 ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Reached the unexpected end of the document", ] );
		$result = 0;
	}
	
	if( wantarray ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Returning the result: $result", "..at depth: $node_depth",
		###LogSD		"..to node named: $node_name", "..and node type: $node_type" ] );
		return( $result, $node_depth, $node_name, $node_type );
	}else{
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Returning the result: $result", ] );
		return $result;
	}
}

sub _reader_init{
	my( $self, $reader ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::XMLReader::_reader_init', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"loading any file specific settings", ] );
	
	if( $self->_get_unique_bits ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"loading any file specific settings - since this is the first open", ] );
		$self->_need_unique_bits( 0 );
		
		# Set basic xml values
		my	$xml_string = '<?xml version="';
		$self->_next_node;
		if( $self->_version ){
			$self->_set_xml_version( $self->_version );
			$xml_string .= $self->_version . '"';
		}else{
			confess "Could not find the version of this xml document!";
		}
		if( $self->_encoding ){
			$self->_set_xml_encoding( $self->_encoding );
			$xml_string = ' encoding="' . $self->_encoding . '"'
		}else{
			$self->_clear_xml_encoding;
		}
		$xml_string .= '?>';
		$self->_set_xml_header( $xml_string );
	
		# Set the file unique bits
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Check if this type of file has unique settings" ], );
		if( $self->can( '_load_unique_bits' ) ){
			###LogSD	$phone->talk( level => 'debug', message =>[ "Loading unique bits" ], );
			$self->_load_unique_bits;
			###LogSD	$phone->talk( level => 'debug', message =>[
			###LogSD		"Finished loading unique bits" 			], );
		}
		$self->start_the_file_over;
	}else{
		###LogSD	$phone->talk( level => 'debug', message =>[ 
		###LogSD		"This is not the first time the file has been opened - don't reload settings" ], );
	}
	###LogSD	$phone->talk( level => 'debug', message => [ "Finished the file initialization" ], );
}

sub _build_xml_reader{
	my( $self, $file_handle ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::XMLReader::_build_xml_reader', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"turning a file handle into an xml reader", ] );
	
	# Build the reader
	$file_handle->seek( 0, 0 );
	my	$xml_reader = XML::LibXML::Reader->new( IO => $file_handle );
		$xml_reader->read;
	###LogSD	$phone->talk( level => 'debug', message =>[ 'Loading reader:', $xml_reader ], );
	$self->_set_xml_parser( $xml_reader );
	###LogSD	$phone->talk( level => 'debug', message => [ "Finished loading XML reader" ], );
}

sub DEMOLISH{
	my ( $self ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::XMLReader::DEMOLISH', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"clearing the reader for log space:" . $self->get_log_space, ] );
	if( $self->_get_xml_parser ){
		#~ print "Disconnecting the sheet file handle from the parser\n";
		###LogSD	$phone->talk( level => 'debug', message =>[ "Disconnecting the file handle from the xml parser", ] );
		$self->_close_the_sheet;
		$self->_clear_xml_parser;
	}
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::XMLReader - A LibXML::Reader xlsx base class

=head1 SYNOPSIS

	package MyPackage;
	use MooseX::StrictConstructor;
	use MooseX::HasDefaults::RO;
	extends	'Spreadsheet::XLSX::Reader::LibXML::XMLReader';
    
=head1 DESCRIPTION

This documentation is written to explain ways to use this module when writing your own excel 
parser.  To use the general package for excel parsing out of the box please review the 
documentation for L<Workbooks|Spreadsheet::XLSX::Reader::LibXML>,
L<Worksheets|Spreadsheet::XLSX::Reader::LibXML::Worksheet>, and 
L<Cells|Spreadsheet::XLSX::Reader::LibXML::Cell>

This module provides a generic way to open an xml file or xml file handle and then extract 
information using the L<XML::LibXML::Reader> parser.  The additional methods and attributes 
are intended to provide some coalated parsing commands that are specifically useful in turning 
xml to perl data structures.

=head2 Attributes

Data passed to new when creating an instance.  For modification of these attributes see the 
listed 'attribute methods'. For general information on attributes see 
L<Moose::Manual::Attributes>.  For ways to manage the instance when opened see the 
L<Methods|/Methods>.
	
=head3 file

=over

B<Definition:> This attribute holds the file handle for the file being read.  If the full 
file name and path is passed to the attribute it is coerced to an IO::File file handle.

B<Default:> no default - this must be provided to read a file

B<Required:> yes

B<Range:> any unencrypted xml file name and path or IO::File file handle

B<attribute methods> Methods provided to adjust this attribute
		
=over

B<set_file>

=over

B<Definition:> change the file value in the attribute (this will reboot 
the file instance and lock the file)

=back

B<get_file>

=over

B<Definition:> Returns the file handle of the file even if a file name 
was passed

=back

B<has_file>

=over

B<Definition:> this is used to see if the file loaded correctly.

=back

B<clear_file>

=over

B<Definition:> this clears (and unlocks) the file handle

=back

=back

L<Delegated Methods>

=over

B<close>

=over

closes the file handle

=back

=back

=back

=head3 error_inst

=over

B<Definition:> This attribute holds the L<error handler
|Spreadsheet::XLSX::Reader::LibXML::Error>.

B<Default:> no default - this must be provided to read a file

B<Required:> yes

B<Range:> any object instance that can provide the required delegated methods.

B<attribute methods> Methods provided to adjust this attribute

=over

B<_clear_error_inst>

=over

clear the attribute value

=back

=back

B<_get_error_inst>

=over

get the attribute value

=back

=back

B<Delegated Methods (required)> Methods delegated to this module by the attribute
		
=over

B<error>

=over

B<Definition:> returns the currently stored error string

=back

B<set_error>

=over

B<Definition:> Sets the error string

=back

B<clear_error>

=over

B<Definition:> clears the error string

=back

B<set_warnings>

=over

B<Definition:> Sets the state that determins if the instance pro-activly 
warns with the error string when the error string is set.

=back

B<if_warn>

=over

B<Definition:> Returns the current state of the state value from 'set_warnings'

=back

=back

=head3 xml_version

=over

B<Definition:> This stores the xml version stored in the xml header.  It is read 
when the file handle is first set in this sheet.

B<Default:> no default - this is auto read from the header

B<Required:> no

B<Range:> xml versions

B<attribute methods> Methods provided to adjust this attribute

=over

B<_clear_xml_version>

=over

clear the attribute value

=back

=back

B<_set_xml_version>

=over

set the attribute value

=back

=back

=head3 xml_encoding

=over

B<Definition:> This stores the data encoding of the xml file from the xml header.  
It is read when the file handle is first set in this sheet.

B<Default:> no default - this is auto read from the header

B<Required:> no

B<Range:> valid xml file encoding

B<attribute methods> Methods provided to adjust this attribute

=over

B<encoding>

=over

get the attribute value

=back

=back

=over

B<has_encoding>

=over

predicate for the attribute value

=back

=back

=over

B<_clear_xml_encoding>

=over

clear the attribute value

=back

=back

B<_set_xml_encoding>

=over

set the attribute value

=back

=back

=head3 xml_header

=over

B<Definition:> This stores the xml header from the xml file.  It is read when 
the file handle is first set in this sheet.

B<Default:> no default - this is auto read from the header

B<Required:> no

B<Range:> valid xml file header

B<attribute methods> Methods provided to adjust this attribute

=over

B<get_header>

=over

get the attribute value

=back

=back

=over

B<_set_xml_header>

=over

set the attribute value

=back

=back

=back

=head3 position_index

=over

B<Definition:> This attribute is available to facilitate other consuming roles and 
classes.  Of the attribute methods only the 'clear_location' method is used in this 
class during the 'start_the_file_over' method.  It can be used for tracking same level 
positions with the same node name.

B<Default:> no default - this is mostly managed by the child class or add on role

B<Required:> no

B<Range:> Integer

B<attribute methods> Methods provided to adjust this attribute

=over

B<where_am_i>

=over

get the attribute value

=back

=back

=over

B<i_am_here>

=over

set the attribute value

=back

=back

=over

B<clear_location>

=over

clear the attribute value

=back

=back

=over

B<has_position>

=over

set the attribute value

=back

=back

=back

=head2 Methods

These are the methods provided by this class.  They most likely should be agumented 
with file specific methods when extending this module.

=head3 start_the_file_over

=over

B<Definition:> This will disconnect the L<XML::LibXML::Reader> from the file handle,  
rewind the file handle, and then reconnect the L<XML::LibXML::Reader> to the file handle.

B<Accepts:> nothing

B<Returns:> nothing

=back

=head3 get_text_node

=over

B<Definition:> This will collect the text node at the current node position.  It will return 
two items ( $success_or_failure, $text_node_value )

B<Accepts:> nothing

B<Returns:> ( $success_or_failure(1|undef), ($text_node_value|undef) )

=back

=head3 get_attribute_hash_ref

=over

B<Definition:> Some nodes have attribute settings.  This method returns a hashref with any 
attribute settings attached as key => value pairs or an empty hash for no attributes

B<Accepts:> nothing

B<Returns:> { attribute_1 => attribute_1_value ... etc. }

=back

=head3 advance_element_position( [$node_name], [$number_of_times_to_index] )

=over

B<Definition:> This method will attempt to advance to $node_name (optional) or the next node 
if no $node_name is passed.  If there is an expectation of multiple nodes of the same name at 
the same level you can also pass $number_of_times_to_index (optional).  This will move through 
the xml file at the $node_name level the number of times indicated starting with wherever the 
xml file is already located.  Meaning $number_of_times_to_index is a relative index not an 
absolute index.

B<Accepts:> nothing

B<Returns:> success or failure for the method call

=back

=head3 location_status

=over

B<Definition:> This method gives three usefull location values with one call

B<Accepts:> nothing

B<Returns:> ( $node_depth (from the top of the file), $node_name, $node_type (xml numerical value for type) );

=back

=head2 Delegated Methods

These are the methods delegated to this class from L<XML::LibXML::Reader>.  For more 
general parsing of subsections of the xml file also see L<Spreadsheet::XLSX::Reader::LibXML>.

=head3 copy_current_node

=over

B<Delegated from:> L<XML::LibXML::Reader/copyCurrentNode (deep)>

Returns an XML::LibXML::Node object

=back

=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader::LibXML/issues
|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

=back

=head1 TODO

=over

B<1.> Nothing currently

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

L<XML::LibXML::Reader>

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

#########1#########2 main pod documentation end   5#########6#########7#########8#########9