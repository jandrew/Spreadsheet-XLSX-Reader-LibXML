package Spreadsheet::XLSX::Reader::LibXML::XMLReader;
use version; our $VERSION = qv('v0.38.16');

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Types::Standard qw(
		Int				HasMethods			Bool
		Num				Str
    );
use XML::LibXML::Reader;
use lib	'../../../../../lib',;
###LogSD	with 'Log::Shiras::LogSpace';
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use Spreadsheet::XLSX::Reader::LibXML::Types qw(
		IOFileType
	);

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
		$self->_clear_location;
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
	if( $self->has_value ){
		my $node_text = $self->node_value;
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
	my $result = $self->move_to_first_att;
	###LogSD	$phone->talk( level => 'trace', message =>[
	###LogSD		"Result of the first attribute move: $result",] );
	ATTRIBUTELIST: while( $result > 0 ){
		my $att_name = $self->node_name;
		my $att_value = $self->node_value;
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
		$result = $self->move_to_next_att;
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
	###LogSD	$phone->talk( level => 'debug', message => [
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
		###LogSD	$phone->talk( level => 'debug', message => [
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
	my ( $node_depth, $node_name, $node_type ) = ( $self->node_depth, $self->node_name, $self->node_type );
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
		#~ byte_consumed		=> 'byteConsumed',
		_read_next_node		=> 'read',
		_next_element		=> 'nextElement',
		node_type			=> 'nodeType',
		node_name			=> 'name',
		node_value			=> 'value',
		has_value			=> 'hasValue',
		node_depth			=> 'depth',
		move_to_first_att	=> 'moveToFirstAttribute',
		move_to_next_att	=> 'moveToNextAttribute',
		_encoding			=> 'encoding',
		_version			=> 'xmlVersion',
		_go_to_the_end		=> 'finish',
		_close_the_sheet	=> 'close',
		#~ get_node_all		=> 'readOuterXml',
	},
	trigger => \&_reader_init,
);

has _position_index =>(
		isa			=> Int,
		reader		=> 'where_am_i',
		writer		=> '_i_am_here',
		clearer		=> '_clear_location',
		predicate	=> 'has_position',
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
	my $result = eval{ $self->_read_next_node } ? 1 : 0 ;
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
			$xml_string .= $self->version . '"';
		}else{
			confess "Could not find the version of this xml document!";
		}
		if( $self->_encoding ){
			$self->_set_xml_encoding( $self->_encoding );
			$xml_string = ' encoding="' . $self->encoding . '"'
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
information using the L<XML::LibXML::Reader> parser.  The class does this by using 
L<delegation|Moose::Manual::Delegation> from the ~::Reader module to import useful functions 
to this module.  Aside from the delegation piece this module provides four other useful 
elements.  First, the module has an attribute to load the file or file handle and uses coercion 
to turn a file into a file handle from the L<Types
|Spreadsheet::XLSX::Reader::LibXML::Types/IOFileType> library.  Second, the module has an 
attribute to store an L<error handler|Spreadsheet::XLSX::Reader::LibXML::Error>.  Third, 
the module provides a L<rewind|/start_the_file_over> function since that is not available in  
the L<XML::LibXML::Reader> parser. Finally, this module has a hook for classes that extend 
this functionality during the initial build.  The initialization of the file will also attempt 
to call '_load_unique_bits'.  It will only call that method once on initialization.

Further use of the module or specialization of the reader can be done by L<extending|/SYNOPSIS> 
the class.

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

B<Range:> any unencrypted xml file name and path or file handle

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

=back

=head3 error_inst

=over

B<Definition:> This attribute holds the L<error handler
|Spreadsheet::XLSX::Reader::LibXML::Error>.

B<Default:> no default - this must be provided to read a file

B<Required:> yes

B<Range:> any object instance that can provide the required delegated methods.

B<delegated methods> Methods provided delegated by the attribute
		
=over

B<error>

=over

B<Definition:> returns the stored error string

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

=back

=head2 Class Methods

These are the methods provided by this class.  They most likely should be agumented 
with file specific methods when extending this module.

=head3 start_the_file_over

=over

B<Definition:> This will disconnect the L<XML::LibXML::Reader> from the file handle,  
rewind the file handle, and then reconnect the L<XML::LibXML::Reader> to the file handle.

B<Accepts:> nothing

B<Returns:> nothing

=back

=head2 Delegated Methods

These are the methods delegated to this class from L<XML::LibXML::Reader>.  For more 
general parsing of subsections of the xml file also see L<Spreadsheet::XLSX::Reader::LibXML>.

=head3 copy_current_node

B<Delegated from:> L<XML::LibXML::Reader/copyCurrentNode (deep)>

=head3 byte_consumed

B<Delegated from:> L<XML::LibXML::Reader/byteConsumed ()>

=head3 start_reading

B<Delegated from:> L<XML::LibXML::Reader/read ()>

=head3 _next_element

B<Delegated from:> L<XML::LibXML::Reader/nextElement>

=head3 node_type

B<Delegated from:> L<XML::LibXML::Reader/nodeType>

=head3 node_name

B<Delegated from:> L<XML::LibXML::Reader/name>

=head3 node_value

B<Delegated from:> L<XML::LibXML::Reader/value>

=head3 has_value

B<Delegated from:> L<XML::LibXML::Reader/hasValue>

=head3 node_depth

B<Delegated from:> L<XML::LibXML::Reader/depth>

=head3 move_to_first_att

B<Delegated from:> L<XML::LibXML::Reader/moveToFirstAttribute>

=head3 move_to_next_att

B<Delegated from:> L<XML::LibXML::Reader/moveToNextAttribute>

=head3 encoding

B<Delegated from:> L<XML::LibXML::Reader/encoding ()>

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