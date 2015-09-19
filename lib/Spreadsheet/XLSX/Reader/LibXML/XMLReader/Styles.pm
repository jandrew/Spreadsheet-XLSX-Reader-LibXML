package Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles;
use version; our $VERSION = qv('v0.38.14');
###LogSD	warn "You uncovered internal logging statements for Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles-$VERSION";

use 5.010;
use Moose;
use MooseX::StrictConstructor;
use MooseX::HasDefaults::RO;
use Carp qw( confess );
use Clone qw( clone );
use Types::Standard qw(
		ArrayRef		HasMethods		Enum
		Bool			Int				is_Int
		is_HashRef
    );
use lib	'../../../../../../lib',;
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
extends	'Spreadsheet::XLSX::Reader::LibXML::XMLReader';

#########1 Dispatch Tables & Package Variables    5#########6#########7#########8#########9

my	$element_lookup ={
		numFmts			=> 'numFmt',
		fonts			=> 'font',
		borders			=> 'border',
		fills			=> 'fill',
		cellStyleXfs	=> 'xf',
		cellXfs			=> 'xf',
		cellStyles		=> 'cellStyle',
		tableStyles		=> 'tableStyle',
	};

my	$key_translations ={
		fontId		=> 'fonts',
		borderId	=> 'borders',
		fillId		=> 'fills',
		xfId		=> 'cellStyles',
		#~ pivotButton	 => 'pivotButton',
	};

my	$cell_attributes ={
		fontId			=> 'cell_font',
		borderId		=> 'cell_border',
		fillId			=> 'cell_fill',
		xfId			=> 'cell_style',
		numFmtId		=> 'cell_coercion',
		alignment		=> 'cell_alignment',
		numFmts			=> 'cell_coercion',
		fonts			=> 'cell_font',
		borders			=> 'cell_border',
		fills			=> 'cell_fill',
		cellStyleXfs	=> 'cellStyleXfs',
		cellXfs			=> 'cellXfs',
		cellStyles		=> 'cell_style',
		tableStyles		=> 'tableStyle',
		#~ pivotButton		=> 'pivotButton',
	};

my	$xml_from_cell ={
		cell_font		=> 'fontId',
		cell_border		=> 'borderId',
		cell_fill		=> 'fillId',
		cell_style		=> 'xfId',
		cell_coercion	=> 'numFmtId',
		cell_alignment	=> 'alignment',
	};

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9
	
has cache_positions =>(
		isa		=> Bool,
		reader	=> '_should_cache_positions',
		default	=> 1,
	);

has format_inst =>(
		isa		=> HasMethods[qw( get_defined_conversion set_defined_excel_formats )],
		handles	=>[qw( get_defined_conversion set_defined_excel_formats )],
	);

has empty_return_type =>(
		isa		=> Enum[qw( empty_string undef_string )],
		reader	=> 'get_empty_return_type',
		writer	=> 'set_empty_return_type',
	);
with	'Spreadsheet::XLSX::Reader::LibXML::XMLToPerlData';#::XMLReader=

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9

sub get_format_position{
	my( $self, $position, $header, $exclude_header ) = @_;
	my	$xml_target_header = $header ? $header : '';#$xml_from_cell->{$header}
	my	$xml_exclude_header = $exclude_header ? $xml_from_cell->{$exclude_header} : '';
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::get_format_position', );
	###LogSD		$phone->talk( level => 'info', message => [
	###LogSD			"Get defined formats at position: $position",
	###LogSD			( $header ? "Returning only the values for header: $header - $xml_target_header" : '' ),
	###LogSD			( $exclude_header ? "..excluding the values for header: $exclude_header - $xml_exclude_header" : '' ) , ] );
	
	# Check for stored value - when caching implemented
	my	$already_got_it = 0;
	if( $self->_has_styles_positions ){
		if( $position > $self->_get_styles_count - 1 ){
			$self->set_error( "Requested styles position is out of range for this workbook" );
			return undef;
		}
		my $target_ref = clone( $self->_get_s_position( $position ) );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"The complete cached style is:", $target_ref, ] );
		if( $header ){
			$target_ref = $target_ref->{$header} ? { $header => $target_ref->{$header} } : undef;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"The cached style with target header only is:", $target_ref  ] );
		}elsif( $exclude_header ){
			delete $target_ref->{$exclude_header};
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"The cached style with exclude header -$exclude_header- removed is:", $target_ref  ] );
		}
		return $target_ref;
	}
	
	# pull the value the long (hard and slow) way
	my ( $node_depth, $node_name, $node_type ) = $self->location_status;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Starting at node: $node_name",  "..at node depth: $node_depth", "..and node type: $node_type"  ] );
	
	# Pull the base ref
	my( $success, $base_ref ) = $self->_get_header_and_value( 'cellXfs', $position );
	if( !$success ){
		confess "Unable to pull position -$position- of the base stored formats (cellXfs)";
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Building out the position -$position- for:", $base_ref,
	###LogSD		( $header ? "..with target header: $xml_target_header" : '' ), ( $exclude_header ? "..and exclude header: $xml_exclude_header" : '' ) ] );
	my $built_ref = $self->_build_perl_style_formats( $base_ref, $xml_target_header, $xml_exclude_header );
	###LogSD	$phone->talk( level => 'trace', message => [
	###LogSD		"Built position -$position- is:", $built_ref ] );
	
	return $built_ref;
}

sub get_default_format_position{
	my( $self, $header, $exclude_header ) = @_;
	my	$position = 0;
	my	$xml_target_header = $header ? $xml_from_cell->{$header} : '';
	my	$xml_exclude_header = $exclude_header ? $xml_from_cell->{$exclude_header} : '';
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::get_default_format_position', );
	###LogSD		$phone->talk( level => 'info', message => [
	###LogSD			"Get defined formats at default position: $position",
	###LogSD			( $header ? "Returning only the values for header: $header - $xml_target_header" : '' ),
	###LogSD			( $exclude_header ? "..excluding the values for header: $exclude_header - $xml_exclude_header" : '' ) , ] );
	
	# Check for stored value - when caching implemented
	my	$already_got_it = 0;
	if( $self->_has_generic_styles_positions ){
		if( $position > $self->_get_generic_styles_count - 1 ){
			$self->set_error( "Requested default styles position is out of range for this workbook" );
			return undef;
		}
		my $target_ref = $self->_get_gs_position( $position );
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"The complete cached style is:", $target_ref  ] );
		if( $header ){
			$target_ref = $target_ref->{$header} ? { $header => $target_ref->{$header} } : undef;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"The cached style with target header only is:", $target_ref  ] );
		}elsif( $exclude_header ){
			delete $target_ref->{$exclude_header};
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"The cached style with exclude header -$exclude_header- removed is:", $target_ref  ] );
		}
		return $target_ref;
	}
	
	# pull the value the long (hard and slow) way
	my ( $node_depth, $node_name, $node_type ) = $self->location_status;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Starting at node: $node_name",  "..at node depth: $node_depth", "..and node type: $node_type"  ] );
	
	# Pull the base ref
	my( $success, $base_ref ) = $self->_get_header_and_value( 'cellStyleXfs', $position );
	if( !$success ){
		confess "Unable to pull position -$position- of the base stored generic formats (cellStylesXfs)";
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Building out the position -$position- for:", $base_ref,
	###LogSD		( $header ? "..with target header: $xml_target_header" : '' ), ( $exclude_header ? "..and exclude header: $xml_exclude_header" : '' ) ] );
	my $built_ref = $self->_build_perl_style_formats( $base_ref, $xml_target_header, $xml_exclude_header );
	###LogSD	$phone->talk( level => 'trace', message => [
	###LogSD		"Built position -$position- is:", $built_ref ] );
	
	return $built_ref;
}

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9
	
has _styles_positions =>(
		isa		=> ArrayRef,
		traits	=> ['Array'],
		handles	=>{
			_get_s_position => 'get',
			_set_s_position => 'set',
			_add_s_position => 'push',
		},
		reader => '_get_all_cache',
		predicate => '_has_styles_positions'
	);
	
has _styles_count =>(
		isa		=> Int,
		default	=> 0,
		reader => '_get_styles_count',
		writer => '_set_styles_count',
	);
	
has _generic_styles_positions =>(
		isa		=> ArrayRef,
		traits	=> ['Array'],
		handles	=>{
			_get_gs_position => 'get',
			_set_gs_position => 'set',
			_add_gs_position => 'push',
		},
		reader => '_get_all_generic_cache',
		predicate => '_has_generic_styles_positions'
	);
	
has _generic_styles_count =>(
		isa		=> Int,
		default	=> 0,
		reader => '_get_generic_styles_count',
		writer => '_set_generic_styles_count',
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

###LogSD	sub BUILD {
###LogSD	    my $self = shift;
###LogSD			$self->set_class_space( 'Styles' );
###LogSD	}

sub _load_unique_bits{
	my( $self, ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::_load_unique_bits', );
	#~ ###LogSD		$phone->talk( level => 'trace', message => [ 'self:', $self ] );
	my ( $node_depth, $node_name, $node_type ) = $self->location_status;
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Arrived at _load_unique_bits pointed to node: $node_name", ] );
	if( $node_name ne 'styleSheet' ){
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD		'The file is not indexed where I want it - resetting the file' ] );
		$self->start_the_file_over;
		$self->advance_element_position( 'styleSheet', 1 );
		( $node_depth, $node_name, $node_type ) = $self->location_status;
		###LogSD		$phone->talk( level => 'debug', message => [
		###LogSD			"Reset and got to node name: $node_name", ] );
	}
	
	# Check for a known format
	if( $node_name ne 'styleSheet' ){
		confess "Can't find the styleSheet node in the xml file / section";
	}
	
	# Initial pull from the xml
	my ( $custom_format_ref, $top_level_ref );
	if( $self->_should_cache_positions ){
		$top_level_ref		= $self->parse_element;
		$custom_format_ref	= $top_level_ref->{numFmts} if exists $top_level_ref->{numFmts};
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD		"Parsing the whole thing for caching:", $top_level_ref ] );
	}else{
		if( $self->advance_element_position( 'numFmts' ) ){
			$custom_format_ref = $self->parse_element;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Pulling the custom number formats only:", $custom_format_ref ] );
		}
	}
	
	# Load the custom formats
	if( $custom_format_ref ){
		my	$translations;
		for my $format ( @{$custom_format_ref->{list}} ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Adding sheet defined translations:", $format ] );
			my	$format_code = $format->{formatCode};
				$format_code =~ s/\\//g;
			$translations->[$format->{numFmtId}] = $format_code;
		}
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		'loaded format positions:', $translations ] );
		$self->set_defined_excel_formats( $translations );
	}
	
	# Cache remaining as needed
	my( $list_to_cache, $count );
	if( $self->_should_cache_positions ){
		###LogSD	$phone->talk( level => 'info', message => [
		###LogSD		"Load the rest of the cache" ] );
		$self->close;# Don't need the file open any more!
		$self->clear_file;
		
		# Build specfic formats
		if( !exists $top_level_ref->{cellXfs} ){
			confess "No base level formats (cellXfs) stored";
		}else{
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Building specific cell formats" ] );
			$self->_set_styles_count( $self->_coalate_perl_style_formats(
				$top_level_ref->{cellXfs}->{list},
				'_add_s_position',
				$top_level_ref
			) );
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		"Final specific caches:", $self->_get_all_cache ] );
		}
		
		
		# Build generic formats
		if( exists $top_level_ref->{cellStyleXfs} ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Building generic cell formats" ] );
			$self->_set_generic_styles_count( $self->_coalate_perl_style_formats(
				$top_level_ref->{cellStyleXfs}->{list},
				'_add_gs_position',
				$top_level_ref
			) );
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		"Final generic caches:", $self->_get_all_generic_cache ] );
		}
	}
	return undef;
}

sub _coalate_perl_style_formats{
	my( $self, $list_ref, $list_method, $top_ref ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::_coalate_perl_style_formats', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Coalating a perl style (cell ready) set of refs from the xml parse with: $list_method", $list_ref, $top_ref ] );
	my $count = 0;
	for my $position ( @$list_ref ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Processing position:", $position ] );
		$count++;
		for my $key ( keys %$position ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Processing key: $key", "..at position: $position->{$key}", ] );
			if( $key eq 'numFmtId' ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Pulling the number conversion for position: $position->{$key}", ] );
				$position->{$cell_attributes->{$key}} = $self->get_defined_conversion( $position->{$key} );
			}elsif( is_HashRef( $position->{$key} ) ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Skipping the key -$key- already comes with embedded settings:", $position->{$key}] );
				next;
			}elsif( $key =~ /(apply|pivotButton|quotePrefix)/ ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Skipping the key: $key", ] );
				next;
			}elsif( !exists $cell_attributes->{$key} ){
				$self->set_error( "Format key -$key- not yet supported by this package" );
				exit 1;
				next;
			}else{
				if( is_Int( $position->{$key} ) ){
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Adding the sub-ref:",  $top_ref->{$key_translations->{$key}}->{list}->[$position->{$key}] ] );
					$position->{$cell_attributes->{$key}} = $top_ref->{$key_translations->{$key}}->{list}->[$position->{$key}];
				}else{
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Translating the key for the sub-ref:",  $position->{$key} ] );
					$position->{$cell_attributes->{$key}} = $position->{$key};
				}
			}
		}
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD		"Final position:", $position ] ); 
		$self->$list_method( $position );
	}
	return $count;
}

sub _build_perl_style_formats{
	my( $self, $base_ref, $target_header, $exclude_header ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::_build_perl_style_formats', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building a perl style (cell ready) ref from the base xml ref", $base_ref,
	###LogSD			( $target_header ? "..returning only header: $target_header" : undef ),
	###LogSD			( $exclude_header ? "..excluding header: $exclude_header" : undef ), ] );
	my $return_ref;
	if( $target_header ){
		if( exists $base_ref->{$xml_from_cell->{$target_header}} ){
			$target_header = $xml_from_cell->{$target_header};
		}
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Processing sub group: $target_header", "..with position: $base_ref->{$target_header}", ] );
		$return_ref = { $self->_get_header_and_value( $target_header, $base_ref->{$target_header} ) };
	}else{
		for my $key ( keys %$base_ref ){
			next if $key eq $exclude_header;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Processing sub group: $key", "..with position: $base_ref->{$key}", ] );
			my( $key, $sub_ref ) = $self->_get_header_and_value( $key, $base_ref->{$key} );
			$return_ref->{$key} = $sub_ref;
		}
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Returning ref:", $return_ref ] );
	return $return_ref;
}

sub _get_header_and_value{
	my( $self, $target_header, $target_position ) = @_;
	$target_header = exists $key_translations->{$target_header} ? $key_translations->{$target_header} : $target_header;
	my $sub_header = exists $element_lookup->{$target_header} ? $element_lookup->{$target_header} : 'dealers_choice';
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD			$self->get_all_space . '::_get_header_and_value', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"getting the ref for target header: $target_header",
	###LogSD			"..with sub header: $sub_header",
	###LogSD			"..and position: $target_position",			] );
	
	
	my( $key, $value );
	if( $target_header =~ /^apply/ ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Found an 'apply flag'",			] );
		( $key, $value ) = ( $target_header, $target_position );
	}elsif( $target_header eq 'numFmtId' ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Pulling the number conversion for position: $target_position", ] );
		( $key, $value ) = ( $cell_attributes->{$target_header}, $self->get_defined_conversion( $target_position ) );
	}elsif( !exists $cell_attributes->{$target_header} ){
		$self->set_error( "Format key -$target_header- not yet supported by this package" );
	}else{
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Reaching into the xml for header -$target_header- position: $target_position", ] );
		if( is_Int( $target_position ) ){
			my ( $node_depth, $node_name, $node_type ) = $self->location_status;
			my $sub_header = $element_lookup->{$target_header};
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"For the super header: $target_header",
			###LogSD		"Accessing the Styles file for position -$target_position- of the header: $sub_header",
			###LogSD		"..currently at a node named: $node_name", "..of node type: $node_type", "..and node depth: $node_depth"] );
			
			# Begin at the beginning
			if( $node_name eq $target_header or $self->advance_element_position( $target_header ) ){# Can't tell which sub position you are at :(
			my ( $node_depth, $node_name, $node_type ) = $self->location_status;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Arrived at: $target_header",
				###LogSD		"..currently at a node named: $node_name", "..of node type: $node_type", "..and node depth: $node_depth" ] );
			}else{
				$self->start_the_file_over;
				if( $self->advance_element_position( $target_header ) ){
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Rewound to: $target_header" ] );
				}else{
					return( undef, undef );
				}
			}
	
			# Index to the indicated sub position
			my $result = $self->advance_element_position( $sub_header, $target_position + 1 );
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Advancing to position -$target_position- gives result: $result" ] );
			if( !$result ){
				$self->set_error( "Requested styles sub position for -$target_header- is not found in this workbook" );
				return( undef, undef );
			}
			
			# Pull the data
			my $base_ref = $self->parse_element;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Pulling data from target header -$target_header- for position -$target_position- gives ref:", $base_ref ] );
			
			( $key, $value ) = ( $cell_attributes->{$target_header}, $base_ref );
		}else{
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"just translating the key for the sub-ref:",  $target_position ] );
			( $key, $value ) = ( $cell_attributes->{$key}, $target_position );
		}
	}
	
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Returning key: $key", "..and value:", $value ] );
	return( $key, $value );
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose;
__PACKAGE__->meta->make_immutable;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles - A LibXML::Reader styles base class

=head1 SYNOPSIS

	#!/usr/bin/env perl
	$|=1;
	use Data::Dumper;
	use MooseX::ShortCut::BuildInstance qw( build_instance );
	use Spreadsheet::XLSX::Reader::LibXML::Error;
	use Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles;

	my $file_instance = build_instance(
	    package      => 'StylesInstance',
	    superclasses => ['Spreadsheet::XLSX::Reader::LibXML::XMLReader::Styles'],
	    file         => 'styles.xml',
	    error_inst   => Spreadsheet::XLSX::Reader::LibXML::Error->new,
	    add_roles_in_sequence => [qw(
	        Spreadsheet::XLSX::Reader::LibXML::FmtDefault
	        Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings
	    )],
	);
	print Dumper( $file_instance->get_format_position( 2 ) );

	#######################################
	# SYNOPSIS Screen Output
	# 01: $VAR1 = {
	# 02:    'applyNumberFormat' => '1',
	# 03:    'fontId' => '0',
	# 04:    'fonts'  => {
	# 05:       'color' => {
	# 06:          'theme' => '1'
	# 07:       },
	# 08:       'sz'     => '11',
	# 09:       'name'   => 'Calibri',
	# 10:       'scheme' => 'minor',
	# 11:       'family' => '2'
	# 12:    },
	# 13:    'numFmtId' => '164',
	# 14:    'fillId'   => '0',
	# 15:    'xfId'     => '0',
	# 16:    'borders' => {
	# 17:       'left'     => 1,
	# 18:       'right'    => 1,
	# 19:       'top'      => 1,
	# 20:       'diagonal' => 1,
	# 21:       'bottom'   => 1
	# 22:    },
	# 23:    'borderId' => '0',
	# 24:    'cellStyleXfs' => {
	# 25:       'fillId'   => '0',
	# 26:       'fontId'   => '0',
	# 27:       'borderId' => '0',
	# 28:       'numFmtId' => '0'
	# 29:    },
	# 30:    'fills' => {
	# 31:       'patternFill' => {
	# 32:          'patternType' => 'none'
	# 33:       }
	# 34:    },
	# 35:    'numFmts' => bless( {
	# 36:       'name' => 'Excel_date_164',
	# 37:       'uniq' => 86,
	# 38:       'coercion' => bless( { 
                    ~ 180 lines hidden ~
	# 219:      }, 'Type::Coercion' )
	# 220:    }, 'Type::Tiny' )
	# 221: };
	#######################################

=head1 DESCRIPTION

This documentation is written to explain ways to use this module.  To use the general 
package for excel parsing out of the box please review the documentation for L<Workbooks
|Spreadsheet::XLSX::Reader::LibXML>, L<Worksheets
|Spreadsheet::XLSX::Reader::LibXML::Worksheet>, and 
L<Cells|Spreadsheet::XLSX::Reader::LibXML::Cell>.

This class is written to get useful data from the sub file 'styles.xml' that is 
a member of a zipped (.xlsx) archive or a stand alone XML text file of the same format.  
The styles.xml file contains the format and display options used by Excel for showing 
the stored data.  To unzip an Excel file manually change the \.xlsx extention to \.zip 
and windows should do (most) of the rest.  For linux use an unzip utility. (
L<Archive::Zip> for instance :)

This documentation is the explanation of this specific module.  For a general explanation 
of the class and how to to add or adjust its place in the larger package see the L<Styles
|Spreadsheet::XLSX::Reader::LibXML::Styles> POD.

This module is the simplified way to extract information from the styles file needed when 
doing high level reading of an Excel spread sheet.  In order to do so it subclasses the module 
L<Spreadsheet::XLSX::Reader::LibXML::XMLReader> and leverages one hard coded role 
L<Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData> Additionally the module will 
error if not built with roles that supply two additional methods.  The methods are 
L<get_defined_excel_format|Spreadsheet::XLSX::Reader::LibXML::FmtDefault/get_defined_excel_format( $integer )> 
and L<parse_excel_format_string
|Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings/parse_excel_format_string( $string )>.  
The links lead to the default source of these methods in the package.  I<These methods are 
intentionally not hard coded to this class so that the user can change them at run time.  See 
the attributes L<Spreadsheet::XLSX::Reader::LibXML/default_format_list> and
L<Spreadsheet::XLSX::Reader::LibXML/format_string_parser> for more explanation.>   Read about 
the function of each when replacing them.  If you want to use the roles as-is, one way to 
integrate them is with L<MooseX::ShortCut::BuildInstance>. The 'on-the-fly' roles also 
add other methods (not documented here) to this class.  Look at the documentation for those 
modules to see what else comes with them.

=head2 Method(s)

These are the methods just provided by this class.  Look at the documentation for the the two 
modules consumed by this class for their elements. L<Spreadsheet::XLSX::Reader::LibXML::XMLReader> 
and L<Spreadsheet::XLSX::Reader::LibXML::XMLReader::XMLToPerlData> 

=head3 get_format_position( $position, [$header], [$exclude_header] )

=over

B<Definition:> This will return the styles information from the identified $position
(Counting from zero).  the target position is usually drawn from the cell data stored in 
the worksheet.  The information is returned as a perl hash ref.  Since the styles 
data is in two tiers it finds all the subtier information for each indicated piece and 
appends them to the hash ref as values for each type key.  If you only want a specific 
branch then you can add the branch $header key and the returned value will only contain 
that leg.

B<Accepts:> $position = an integer for the styles $position. (required at position 0)

B<Accepts:> $header = the target header key (optional at postion 1) (use the 
L<Spreadsheet::XLSX::Reader::LibXML::Cell/Attributes> that are cell formats as the definition 
of range for this

B<Accepts:> $exclude_header = the target header key (optional at position 2) (use the 
L<Spreadsheet::XLSX::Reader::LibXML::Cell/Attributes> that are cell formats as the definition 
of range for this)

B<Returns:> a hash ref of data

=back

=head3 get_default_format_position( [$header], [$exclude_header] )

=over

B<Definition:> For any cell that does not have a unquely identified format excel generally 
stores a default format for the remainder of the sheet.  This will return the two 
tiered default styles information.  If you only want the default from a specific header 
then add the $header string to the method call.  The information is returned as a perl 
hash ref.

B<Accepts:> $header = the target header key (optional at postion 0) (use the 
L<Spreadsheet::XLSX::Reader::LibXML::Cell/Attributes> that are cell formats as the definition 
of range for this

B<Accepts:> $exclude_header = the target header key (optional at position 1) (use the 
L<Spreadsheet::XLSX::Reader::LibXML::Cell/Attributes> that are cell formats as the definition 
of range for this)

B<Returns:> a hash ref of data

=back

=head1 SUPPORT

=over

L<github Spreadsheet::XLSX::Reader::LibXML/issues
|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

=back

=head1 TODO

=over

B<2.> This was one of the first XMLReader parsers I wrote and the XML parsing is crufty (needs a scrub)

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

#########1#########2 main pod documentation end   5#########6#########7#########8#########9