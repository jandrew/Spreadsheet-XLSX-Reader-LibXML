package Spreadsheet::XLSX::Reader::LibXML::XMLReader::ExtractFile;
use version; our $VERSION = version->declare('v0.40.2');
###LogSD	warn "You uncovered internal logging statements for Spreadsheet::XLSX::Reader::LibXML::XMLReader::ExtractFile-$VERSION";

use	Moose::Role;
use Clone 'clone';
use Carp 'confess';
use Data::Dumper;
use Types::Standard qw( is_HashRef HashRef Str );
requires qw(
	get_header 			start_the_file_over				advance_element_position
	parse_element		set_exclude_match
),
###LogSD	'get_all_space'
;
use File::Temp qw/ :seekable /;
use lib	'../../../../../lib',;
#~ ###LogSD	use Log::Shiras::Telephone;

#########1 Dispatch Tables    3#########4#########5#########6#########7#########8#########9

my	$default_top_level_attributes ={
		xmlns => "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
		'xmlns:r' => "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
	};

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9



#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub extract_file{
    my ( $self, $ref ) = ( @_ );
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::extract_file', );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			'Arrived at extract_file for the workbook general settings:', $ref ] );
	
	# Get the header
	my 	$file_string = $self->get_header;
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		"Header string: $file_string" ] );
	
	# Build a temp file and load it with the file string
	my	$fh = File::Temp->new();
	print $fh $file_string;
	$fh->seek( 0, 0 );
	###LogSD	$self->_print_current_file( $fh );
	
	# impement the passed args
	my ( $method, @args ) = @$ref;
	$self->$method( $fh, @args );
	
	###LogSD	$phone->talk( level => 'debug', message =>[
	###LogSD		'File handle:', $fh ] );
	###LogSD	$self->_print_current_file( $fh );
	return $fh;
}

sub get_headers{
	my( $self, $fh ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::get_headers', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Turning the save header ref from perl data to an xml string" ] );
	
	$self->_perl_ref_to_xml( $fh, $self->_get_file_headers );
}

sub empty_file{
	my( $self, $fh, $empty_header ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::empty_file', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Arrived at empty_file - making sure the document has : $empty_header" ] );
	print $fh "<$empty_header/>";
}

sub get_whole_node{
	my( $self, $fh, @node_list ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::get_whole_node', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Arrived at get_whole_node - attempting to pull the whole node from one of:", @node_list ] );
	my $found_a_node = 0;
	for my $node ( @node_list ){
		$self->start_the_file_over;
		my $result = $self->advance_element_position( $node );
		if( $result ){
			print $fh $self->get_node_all;
			###LogSD	$self->_print_current_file( $fh );
			$found_a_node = 1;
			last;
		}
	}
	if( !$found_a_node ){
		print $fh "<$node_list[-1]/>";
		###LogSD	$self->_print_current_file( $fh );
	}
}
	

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9

has _file_headers =>(
		isa		=> HashRef,
		reader	=> '_get_file_headers',
		writer	=> '_set_file_headers',
		clearer	=> '_clear_file_headers',
	); 

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _load_unique_bits{
	my( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::_hidden::_load_unique_bits', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Setting the WorkbookFileInterface unique bits for the XMLReader" ] );
	
	$self->_clear_file_headers;
	$self->start_the_file_over;
	my $result = $self->advance_element_position( 'Workbook' );
	if( $result ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Found the Workbook section" ] );
		$self->set_exclude_match( '(Table|DocumentProperties|Styles)' );
		my $old_ref = $self->parse_element( 2 );# Just get the headers
		my $new_ref ={
			list =>[ 
				{
					attributes => { %{$old_ref->{attributes}}, %$default_top_level_attributes },
					list =>[],
					list_keys =>[],
				},
			],
			list_keys =>[ 'workbook' ], 
		};
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD		"Initial Workbook headers are:", $new_ref, $old_ref ] );
		
		# Merge sheets and chartsheets to one sub sheet node
		my $x = 0;
		my $node_position_ref;
		for my $node ( @{$old_ref->{list_keys}} ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Processing node: $node", ] );
			my $sub_ref;
			if( $node =~ /sheet/i ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Found a worksheet or chartsheet", ] );
				$node_position_ref->{$node}++;
				$sub_ref = $old_ref->{list}->[$x]->{attributes};
				$sub_ref->{sheetId} //=  $x + 1;
				$sub_ref->{'r:id'} //= 'rId' . ($x + 1);
				$sub_ref->{sheet_type} //= lc( $node );
				$sub_ref->{node_name} //= $node;
				$sub_ref->{node_position} = $node_position_ref->{$node};
				
				# Log the sheet
				$new_ref->{list}->[0]->{list_keys}->[0] = 'sheets';
				$new_ref->{list}->[0]->{list}->[0]->{list_keys}->[$x] = 'sheet';
				$new_ref->{list}->[0]->{list}->[0]->{list}->[$x]->{attributes} = $sub_ref;
				$x++;
			}
		}
		
		# Check for sheets
		if( $new_ref->{list}->[0]->{list_keys}->[0] ne 'sheets' ){
			$self->set_error( "The Workbook node has no sheets" );
			$self->DEMOLISH;# Nuke it - file load fail
		}else{
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Updated Workbook headers are:", $new_ref ] );
			$self->_set_file_headers( $new_ref );
		}
	}else{
		$self->set_error( "Unable to find the 'Workbook' xml node - this may not be a SpreadsheetML document" );
		$self->DEMOLISH;# Nuke it - file load fail
	}
}

sub _perl_ref_to_xml{
	my( $self, $fh, $ref, $is_middle ) = @_;#, $xml_string
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::_hidden::_perl_ref_to_xml', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Turning the ref:", $ref,
	###LogSD			'..into xml for file handle:', $fh,
	###LogSD			( $is_middle ? "..in the middle of a ref" : undef ) ] );
	
	my $link_keys;
	my $need_end_tag = 0;
	if( is_HashRef( $ref ) ){
		
		# Handle attribute refs
		if( exists $ref->{attributes} ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Loading the attribute ref:", $ref->{attributes}  ] );
			print $fh $self->_ref_to_string( $ref->{attributes} );
			###LogSD	$self->_print_current_file( $fh );
		}
		
		# Handle list nodes
		if( exists $ref->{list_keys} ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Processing node list" ] );
			my $x = 0;
			if( $is_middle ){
				print $fh ">";
				###LogSD	$self->_print_current_file( $fh );
			}
			for my $key ( @{$ref->{list_keys}} ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Processing list key: $key" ] );
				print $fh "<$key";
				###LogSD	$self->_print_current_file( $fh );
				$is_middle = 1;
				if( defined $ref->{list}->[$x] ){
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"This node has deeper content - diving down" ] );
					$self->_perl_ref_to_xml( $fh, $ref->{list}->[$x], 1 );
					if( exists $ref->{list}->[$x]->{list_keys} ){
						###LogSD	$phone->talk( level => 'debug', message => [
						###LogSD		"After (sub) list content adding closing node: $key" ] );
						print $fh "</$key>";
					}else{
						###LogSD	$phone->talk( level => 'debug', message => [
						###LogSD		"Probably only an attribute ref down there - closing the node as self contained" ] );
						print $fh "/>";
					}
					###LogSD	$self->_print_current_file( $fh );
				}else{
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"This node has no content closing the tag as self contained" ] );
					print $fh "/>";
					###LogSD	$self->_print_current_file( $fh );
				}
				$x++;
			}
		}
		
	}elsif( !ref $ref ){
		###LogSD	$phone->talk( level => 'warn', message => [
		###LogSD		"This node terminates with a string - adding it and closing it as self contained" ] );
		print $fh '="' . ($ref//'undef') . '"/>';
		###LogSD	$self->_print_current_file( $fh );
	}else{
		confess "I don't know how to handle: " . Dumper( $ref );
	}
	###LogSD	$self->_print_current_file( $fh );
}

sub _ref_to_string{# Currently only works for single level refs
    my ( $self, $ref ) = ( @_ );#, $exclude
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::_hidden::_ref_to_string', );
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			'Making a string out of a (single level) ref:', $ref,]);
	my $string;
	if( is_HashRef( $ref ) ){
		for my $key ( keys %$ref ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Processing key: $key" ] );
			$string .= " $key=" . '"' . $ref->{$key} . '"';
			###LogSD	$phone->talk( level => 'trace', message =>[
			###LogSD		"updated string: ", $string ] );
		}
	}else{
		confess "I was looking for a hash ref but I got ref type: " . (ref( $ref )//'undef');
	}
		
	###LogSD	$phone->talk( level => 'trace', message =>[
	###LogSD		"returning string: ", $string ] );
	return $string;
}

###LogSD	sub _print_current_file{
###LogSD	    my ( $self, $ref ) = ( @_ );
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::_hidden::_print_current_file', );
	###LogSD	my $line =  ( caller(0) )[2];
	###LogSD	$ref->seek( 0, 0 );
	###LogSD	my $next_line;
	###LogSD	while( $next_line = <$ref> ){
	###LogSD		$phone->talk( level => 'debug', message =>[
	###LogSD			"For code line -$line- the next line of the file is:", $next_line ]);
###LogSD	    }
###LogSD	}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose::Role;
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::XMLReader::ExtractFile - XMLReader file extractor

=head1 DESCRIPTION

Not written yet!

=head1 SEE ALSO

=over

L<Spreadsheet::ParseExcel> - Excel 2003 and earlier

L<Spreadsheet::ParseXLSX> - 2007+

L<Spreadsheet::Read> - Generic

L<Spreadsheet::XLSX> - 2007+

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=over

All lines in this package that use Log::Shiras are commented out

=back

=back

=cut

#########1#########2 main pod documentation end  5#########6#########7#########8#########9