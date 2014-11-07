package Spreadsheet::XLSX::Reader::LibXML::ParseExcelFormatStrings;
use version; our $VERSION = qv('v0.5_1');

use 5.010;
use Moose::Role;
requires qw(
	get_log_space				_add_integer_separator
	_continuous_fraction		change_output_encoding
	get_excel_region
);
use Types::Standard qw(
		Int						Str				
		Maybe					Num					
		HashRef					ArrayRef
		CodeRef					Object			
		ConsumerOf				InstanceOf			
		HasMethods				Bool
		is_Object
    );
use Carp qw( confess );
use	Type::Coercion;
use	Type::Tiny;
use DateTimeX::Format::Excel v0.12;
use DateTime::Format::Flexible;
use lib	'../../../../../lib',;
###LogSD	use Log::Shiras::Telephone;
###LogSD	use Log::Shiras::UnhideDebug;
use	Spreadsheet::XLSX::Reader::LibXML::Types v0.5 qw(
		PositiveNum				NegativeNum
		ZeroOrUndef				NotNegativeNum
		Excel_number_0
	);

#########1 Dispatch Tables & Package Variables    5#########6#########7#########8#########9

my $coercion_index		= 0;
my @type_list			= ( PositiveNum, NegativeNum, ZeroOrUndef, Str );
my $last_date_cldr		= 'yyyy-m-d';
my $last_duration		= 0;
my $last_sub_seconds	= 0;
my $last_format_rem		= 0;
my $duration_order		={ h => 'm', m =>'s', s =>'0' };

#########1 Public Attributes  3#########4#########5#########6#########7#########8#########9

has	epoch_year =>(
		isa		=> Int,
		reader	=> 'get_epoch_year',
		default	=> 1900,
	);
	
has	cache_formats =>(
		isa		=> Bool,
		reader	=> 'get_cache_behavior',
		default	=> 1,
	);
	
has	datetime_dates =>(
		isa		=> Bool,
		reader	=> 'get_date_behavior',
		writer	=> 'set_date_behavior',
		default	=> 0,
	);

#########1 Public Methods     3#########4#########5#########6#########7#########8#########9
 
sub parse_excel_format_string{# Currently only handles dates and times
	my( $self, $format_strings, $coercion_name ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new(
	###LogSD			name_space 	=> $self->get_log_space .  '::parse_excel_format_string', );
	if( !defined $format_strings ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Nothing passed to convert",] );
		return Excel_number_0;
	}
	$format_strings =~ s/\\//g;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"parsing the custom excel format string: $format_strings",] );
	my $conversion_type = 'number';
	# Check the cache
	my	$cache_key;
	if( $self->get_cache_behavior ){
		$cache_key	= $format_strings; # TODO fix the non-hashkey character issues;
		if( $self->has_cached_format( $cache_key ) ){
			###LogSD		$phone->talk( level => 'debug', message => [
			###LogSD			"Format already built - returning stored value for: $cache_key", ] );
			return $self->get_cached_format( $cache_key );
		}else{
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Building new format for key: $cache_key", ] );
		}
	}
	
	# Split into the four sections positive, negative, zero, and text
		$format_strings =~ s/General/\@/ig;# Change General to text input
	my	@format_string_list = split /;/, $format_strings;
	my	$last_is_text = ( $format_string_list[-1] =~ /\@/ ) ? 1 : 0 ;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Is the last position text: $last_is_text",	] );
	# Make sure the full range of number inputs are sent down the right path;
	my	@used_type_list = @{\@type_list};
		$used_type_list[0] =
			( scalar( @format_string_list ) - $last_is_text == 1 ) ? Maybe[Num] :
			( scalar( @format_string_list ) - $last_is_text == 2 ) ? Maybe[NotNegativeNum] : $type_list[0] ;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Now operating on each format string", @format_string_list,
	###LogSD		'..with used type list:', map{ $_->name } @used_type_list,	] );
	my	$format_position = 0;
	my	@coercion_list;
	my	$action_type;
	my	$is_date = 0;
	my	$date_text = 0;
	for my $format_string ( @format_string_list ){
		$format_string =~ s/_.//g;# no character justification to other rows
		$format_string =~ s/\*//g;# Remove the repeat character listing (not supported here)
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Building format for: $format_string", ] );
		
		# Pull out all the straight through stuff
		my @deconstructed_list;
		my $x = 0;
		while( defined $format_string and my @result = $format_string =~
					/^(													# Collect any formatting stuff first
						(AM\/PM|										# Date 12 hr flag
						A\/P|											# Another date 12 hr flag
						\[hh?\]|										# Elapsed hours
						\[mm\]|											# Elapsed minutes
						\[ss\]|											# Elapsed seconds
						[dmyhms]+)|										# DateTime chunks
						([0-9#\?]+[,\-\_]?[#0\?]*,*|					# Number string
						\.|												# Split integers from decimals
						[Ee][+\-]|										# Exponential notiation
						%)|												# Percentage
						(\@)											# Text input
					)?(													# Finish collecting format actions
					(\"[^\"]*\")|										# Anything in quotes just passes through
					(\[[^\]]*\])|										# Anything in brackets needs modification
					[\(\)\$\-\+\/\:\!\^\&\'\~\{\}\<\>\=\s]|				# All the pass through characters
					\,\s												# comma space for verbal separation
					)?(.*)/x											){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Now processing: $format_string", '..with result:', @result ] );
			my	$pre_action		= $1;
			my	$date			= $2;
			my	$number			= $3;
			my	$text			= $4;
			my	$fixed_value	= $5;
				$format_string	= $8;
			if( $fixed_value ){
				if( $fixed_value =~ /\[\$([^\-\]]*)\-?\d*\]/ ){# removed the localized element of fixed values
					$fixed_value = $1;
				}elsif( $fixed_value =~ /\[[^hms]*\]/ ){# Remove all color and conditionals as they will not be used
					$fixed_value = undef;
				}elsif( $fixed_value =~ /\"\-\"/ and $format_string ){# remove decimal justification for zero bars
					###LogSD	$phone->talk( level => 'trace', message => [
					###LogSD		"Initial format string: $format_string", ] );
					$format_string =~ s/^(\?+)//;
					###LogSD	$phone->talk( level => 'trace', message => [
					###LogSD		"updated format string: $format_string", ] );
				}
			}
			if( defined $pre_action ){
				my	$current_action =
						( $date ) ? 'DATE' :
						( defined $number ) ? 'NUMBER' :
						( $text ) ? 'TEXT' : 'BAD' ;
					$is_date = 1 if $date;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Current action from -$pre_action- is: $current_action" ] );
				if( $action_type and $current_action and ($current_action ne $action_type) ){
					###LogSD	$phone->talk( level => 'info', message => [
					###LogSD		"General action type: $action_type",,
					###LogSD		"is failing current action: $current_action", ] );
					my $fail = 1;
					if( $action_type eq 'DATE' ){
						$conversion_type = 'date';
						###LogSD	$phone->talk( level => 'info', message => [
						###LogSD		"Checking the date mishmash", ] );
						if( $current_action eq 'NUMBER' ){
							###LogSD	$phone->talk( level => 'info', message => [
							###LogSD		"Special case of number following action", ] );
							if(	( $pre_action =~ /^\.$/ and $format_string =~ /^0+/				) or
								( $pre_action =~ /^0+$/ and $deconstructed_list[-1]->[0] =~ /^\.$/	)	){
								$current_action = 'DATE';
								$fail = 0;
							}
						}elsif( $pre_action eq '@' ){
							###LogSD	$phone->talk( level => 'info', message => [
							###LogSD		"Excel conversion of pre-epoch datestring pass through highjacked here", ] );
							$current_action = 'DATESTRING';
							$fail = 0;
						}
					}elsif( $action_type eq 'NUMBER' ){
						###LogSD	$phone->talk( level => 'info', message => [
						###LogSD		"Checking for possible text in a number field for a pass throug", ] );
						if( $current_action eq 'TEXT' ){
							###LogSD	$phone->talk( level => 'info', message => [
							###LogSD		"Special case of text following a number", ] );
							$fail = 0;
						}
					}
					if( $fail ){
						confess "Bad combination of actions in this format string: $format_strings - $action_type - $current_action";
					}
				}
				$action_type = $current_action if $current_action;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		(($pre_action) ? "First action resolved to: $pre_action" : undef),
				###LogSD		(($fixed_value) ? "Extracted fixed value: $fixed_value" : undef),
				###LogSD		(($format_string) ? "Remaining string: $format_string" : undef),
				###LogSD		"With updated deconstruction list:", @deconstructed_list, ] );
			}else{
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Early elements unusable - remaining string: $format_string", ] );
			}
			push @deconstructed_list, [ $pre_action, $fixed_value ];
			if( $x++ == 30 ){
				confess "Regex matching failed (with an infinite loop) for excel format string: $format_string";
			}
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		(($pre_action) ? "First action resolved to: $pre_action" : undef),
			###LogSD		(($fixed_value) ? "Extracted fixed value: $fixed_value" : undef),
			###LogSD		(($format_string) ? "Remaining string: $format_string" : undef),
			###LogSD		"With updated deconstruction list:", @deconstructed_list, ] );
			last if length( $format_string ) == 0;
		}
		push @deconstructed_list, [ $format_string, undef ] if $format_string;
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Current segment type: $action_type", 
		###LogSD		"List with fixed values separated:", @deconstructed_list ] );
		my $method = '_build_' . lc($action_type);
		my $filter = ( $action_type eq 'TEXT' ) ? Str : $used_type_list[$format_position++];
		if( $action_type eq 'DATESTRING' ){
			$date_text = 1;
			$filter = Str;
		}
		push @coercion_list, $self->$method( $filter, \@deconstructed_list );
	}
	push @coercion_list, $self->_build_datestring( Str, [ [ '@', '' ] ] ) if $is_date and !$date_text;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		'Length of coersion list: ' . scalar( @coercion_list ),
	###LogSD		(map{ if( is_Object( $_ ) and $_->can( 'name' ) ){ $_->name }else{ $_ } } @coercion_list), ] );
	
	# Build the final format
	$coercion_name =~ s/__/_${conversion_type}_/ if $coercion_name;
	my	%args = (
			name		=> ($coercion_name // ($action_type . '_' . $coercion_index++)),
			coercion	=>	Type::Coercion->new(
								type_coercion_map => [ @coercion_list ],
							),
			#~ coerce		=> 1,
		);
	my	$final_type = Type::Tiny->new( %args );
	
	# Save the cache
	$self->set_cached_format( $cache_key => $final_type ) if $self->get_cache_behavior;
	
	return $final_type;
}
	

#########1 Private Attributes 3#########4#########5#########6#########7#########8#########9
	
has	_format_cash =>(
		isa		=> HashRef,
		traits	=> ['Hash'],
		handles =>{
			has_cached_format => 'exists',
			get_cached_format => 'get',
			set_cached_format => 'set',
		},
		default	=> sub{ {} },
	);

#########1 Private Methods    3#########4#########5#########6#########7#########8#########9

sub _build_text{
	my( $self, $type_filter, $list_ref ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_text', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building an anonymous sub to process text values" ] );
	my $sprintf_string;
	my $found_string = 0;
	for my $piece ( @$list_ref ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"processing text piece:", $piece ] );
		if( !$found_string and defined $piece->[0] ){
			$sprintf_string .= '%s';
			$found_string = 1;
		}
		if( $piece->[1] ){
			$sprintf_string .= $piece->[1];
		}
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Final sprintf string: $sprintf_string" ] );
	my	$return_sub = sub{
			my $adjusted_input = $self->change_output_encoding( $_[0] );
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Updated Input: $adjusted_input" ] );
			return sprintf( $sprintf_string, $adjusted_input );
		};
	return( Str, $return_sub );
}

sub _build_date{
	my( $self, $type_filter, $list_ref ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_date', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building an anonymous sub to process date values" ] );
	
	my ( $cldr_string, $format_remainder );
	my	$is_duration = 0;
	my	$sub_seconds = 0;
	if( !$self->get_date_behavior ){
		# Process once to build the cldr string
		my $prior_duration;
		for my $piece ( @$list_ref ){
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"processing date piece:", $piece ] );
			if( defined $piece->[0] ){
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"Manageing the cldr part: " . $piece->[0] ] );
				if( $piece->[0] =~ /\[(.+)\]/ ){
					###LogSD	$phone->talk( level => 'debug', message =>[ "Possible duration" ] );
					(my $initial,) = split //, $1;
					my $length = length( $1 );
					$is_duration = [ $initial, 0, [ $piece->[1] ], [ $length ] ];
					if( $is_duration->[0] =~ /[hms]/ ){
						$piece->[0] = '';
						$piece->[1] = '';
						$prior_duration = 	$is_duration->[0];
						###LogSD	$phone->talk( level => 'debug', message => [
						###LogSD		"found a duration piece:", $is_duration,
						###LogSD		"with prior duration: $prior_duration"		] );
					}else{
						confess "Bad duration element found: $is_duration->[0]";
					} 
				}elsif( ref( $is_duration ) eq 'ARRAY' ){
					###LogSD	$phone->talk( level => 'debug', message =>[ "adding to duration", $piece ] );
					my	$next_duration = $duration_order->{$prior_duration};
					if( $piece->[0] eq '.' ){
						push @{$is_duration->[2]}, $piece->[0];
						###LogSD	$phone->talk( level => 'debug', message => [
						###LogSD		"found a period" ] );
					}elsif( $piece->[0] =~ /$next_duration/ ){
						my $length = length( $piece->[0] );
						$is_duration->[1]++;
						push @{$is_duration->[2]}, $piece->[1] if $piece->[1];
						push @{$is_duration->[3]}, $length;
						($prior_duration,) = split //, $piece->[0];
						if( $piece->[0] =~ /^0+$/ ){
							$piece->[0] =~ s/0/S/g;
							$sub_seconds = $piece->[0];
							###LogSD	$phone->talk( level => 'debug', message => [
							###LogSD		"found a subseconds format piece: $sub_seconds" ] );
						}
						$piece->[0] = '';
						$piece->[1] = '';
						###LogSD	$phone->talk( level => 'debug', message => [
						###LogSD		"Current duration:", $is_duration,
						###LogSD		"with prior duration: $prior_duration"	 ] );
					}else{
						confess "Bad duration element found: $piece->[0]";
					} 
				}elsif( $piece->[0] =~ /m/ ){
					###LogSD	$phone->talk( level => 'debug', message =>[ "Minutes or Months" ] );
					if( ($cldr_string and $cldr_string =~ /:'?$/) or ($piece->[1] and $piece->[1] eq ':') ){
						###LogSD	$phone->talk( level => 'debug', message => [
						###LogSD		"Found minutes - leave them alone" ] );
					}else{
						$piece->[0] =~ s/m/L/g;
						###LogSD	$phone->talk( level => 'debug', message => [
						###LogSD		"Converting to cldr stand alone months (m->L)" ] );
					}
				}elsif( $piece->[0] =~ /h/ ){
					$piece->[0] =~ s/h/H/g;
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Converting 12 hour clock to 24 hour clock" ] );
				}elsif( $piece->[0] =~ /AM?\/PM?/i ){
					$cldr_string =~ s/H/h/g;
					$piece->[0] = 'a';
					###LogSD	$phone->talk( level => 'debug', message =>[ "Set 12 hour clock and AM/PM" ] );
				}elsif( $piece->[0] =~ /d{3,5}/ ){
					$piece->[0] =~ s/d/E/g;
					###LogSD	$phone->talk( level => 'debug', message =>[ "Found a weekday request" ] );
				}elsif( !$sub_seconds and $piece->[0] =~ /[\.]/){#
					$piece->[0] = "'.'";
					#~ $piece->[0] = "':'";
					$sub_seconds = 1;
					###LogSD	$phone->talk( level => 'debug', message =>[ "Starting sub seconds" ] );
				}elsif( $sub_seconds eq '1' ){
					###LogSD	$phone->talk( level => 'debug', message =>[ "Formatting sub seconds" ] );
					if( $piece->[0] =~ /^0+$/ ){
						$piece->[0] =~ s/0/S/g;
						$sub_seconds = $piece->[0];
						$piece->[0] = '';
						###LogSD	$phone->talk( level => 'debug', message => [
						###LogSD		"found a subseconds format piece: $sub_seconds" ] );
					}else{
						confess "Bad sub-seconds element after [$cldr_string] found: $piece->[0]";
					}
				}
				if( $sub_seconds and $sub_seconds ne '1' ){
					$format_remainder .= $piece->[0];
				}else{
					$cldr_string .= $piece->[0];
				}
			}
			if( $piece->[1] ){
				if( $sub_seconds and $sub_seconds ne '1' ){
					$format_remainder .= $piece->[1];
				}else{
					$cldr_string .= $piece->[1];
				}
			}
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		(($cldr_string) ? "Updated CLDR string: $cldr_string" : undef),
			###LogSD		(($format_remainder) ? "Updated format remainder: $format_remainder" : undef),
			###LogSD		(($is_duration) ? ('Duration ref:', $is_duration) : undef)			] );
		}
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"Updated CLDR string: $cldr_string",
		###LogSD		(($is_duration) ? ('...and duration:', $is_duration) : undef )	] );
		$last_date_cldr 	= $cldr_string;
		$last_duration		= $is_duration;
		$last_sub_seconds	= $sub_seconds;
		$last_format_rem	= $format_remainder;
	}
	my	@args_list = ( $self->get_epoch_year == 1904 ) ? ( system_type => 'apple_excel' ) : ();
	my	$converter = DateTimeX::Format::Excel->new( @args_list );
	my	$conversion_sub = sub{ 
			my	$num = $_[0];
			if( !defined $num ){
				return undef;
			}
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Processing date number: $num",
			###LogSD		'..with duration:', $is_duration,
			###LogSD		"..and sub-seconds: $sub_seconds",
			###LogSD		(($format_remainder) ? "..and format_remainder: $format_remainder" : undef) ] );
			my	$dt = $converter->parse_datetime( $num );
			my $return_string;
			my $calc_sub_secs;
			if( $is_duration ){
				my	$di = $dt->subtract_datetime_absolute( $converter->_get_epoch_start );
				if( $self->get_date_behavior ){
					return $di;
				}
				my	$sign = DateTime->compare_ignore_floating( $dt, $converter->_get_epoch_start );
				$return_string = ( $sign == -1 ) ? '-' : '' ;
				my $key = $is_duration->[0];
				my $delta_seconds	= $di->seconds;
				my $delta_nanosecs	= $di->nanoseconds;
				$return_string .= $self->_build_duration( $is_duration, $delta_seconds, $delta_nanosecs );
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Duration return string: $return_string" ] );
			}else{
				if( $self->get_date_behavior ){
					return $dt;
				}
				if( $sub_seconds ){
					$calc_sub_secs = $dt->format_cldr( $sub_seconds );
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Processing sub-seconds: $calc_sub_secs" ] );
					if( "0.$calc_sub_secs" >= 0.5 ){
						###LogSD	$phone->talk( level => 'debug', message => [
						###LogSD		"Rounding seconds back down" ] );
						$dt->subtract( seconds => 1 );
					}
				}
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Converting it with CLDR string: $cldr_string" ] );
				$return_string .= $dt->format_cldr( $cldr_string );
				if( $sub_seconds and $sub_seconds ne '1' ){
					$return_string .= $calc_sub_secs;
				}
				$return_string .= $dt->format_cldr( $format_remainder ) if $format_remainder;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"returning: $return_string" ] );
			}
			return $return_string;
		};
	return( $type_filter, $conversion_sub );
}

sub _build_datestring{
	my( $self, $type_filter, $list_ref ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_datestring', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building an anonymous sub to process date strings" ] );
	
	my ( $cldr_string, $format_remainder );
	#~ my	$is_duration = 0;
	#~ if( !$self->get_date_behavior ){
		#~ $cldr_string = $last_date_cldr;
	#~ }
	my	$conversion_sub = sub{ 
			my	$date = $_[0];
			if( !$date ){
				return undef;
			}
			my $calc_sub_secs;
			if( $date =~ /(.*:\d+)\.(\d+)(.*)/ ){
				$calc_sub_secs = $2;
				$date = $1;
				$date .= $3 if $3;
				$calc_sub_secs .= 0 x (9 - length( $calc_sub_secs ));
			}
			my	$dt = 	DateTime::Format::Flexible->parse_datetime(
							$date, lang =>[ $self->get_excel_region ]
						);
			$dt->add( nanoseconds => $calc_sub_secs ) if $calc_sub_secs;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Processing date string: $date",
			###LogSD		"..with duration:", $last_duration,
			###LogSD		"..and sub-seconds: $last_sub_seconds",
			###LogSD		"..and stripped nanoseconds: $calc_sub_secs"		] );
			my $return_string;
			if( $last_duration ){
				my	@args_list = ( $self->get_epoch_year == 1904 ) ? ( system_type => 'apple_excel' ) : ();
				my	$converter = DateTimeX::Format::Excel->new( @args_list );
				my	$di = $dt->subtract_datetime_absolute( $converter->_get_epoch_start );
				if( $self->get_date_behavior ){
					return $di;
				}
				my	$sign = DateTime->compare_ignore_floating( $dt, $converter->_get_epoch_start );
				$return_string = ( $sign == -1 ) ? '-' : '' ;
				my $key = $last_duration->[0];
				my $delta_seconds	= $di->seconds;
				my $delta_nanosecs	= $di->nanoseconds;;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Delta seconds: $delta_seconds",
				###LogSD		(($delta_nanosecs) ? "Delta nanoseconds: $delta_nanosecs" : undef) ] );
				$return_string .= $self->_build_duration( $last_duration, $delta_seconds, $delta_nanosecs );
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Duration return string: $return_string" ] );
			}else{
				if( $self->get_date_behavior ){
					return $dt;
				}
				if( $last_sub_seconds ){
					$calc_sub_secs = $dt->format_cldr( $last_sub_seconds );
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Processing sub-seconds: $calc_sub_secs" ] );
					if( "0.$calc_sub_secs" >= 0.5 ){
						###LogSD	$phone->talk( level => 'debug', message => [
						###LogSD		"Rounding seconds back down" ] );
						$dt->subtract( seconds => 1 );
					}
				}
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Converting it with CLDR string: $last_date_cldr" ] );
				$return_string .= $dt->format_cldr( $last_date_cldr );
				if( $last_sub_seconds and $last_sub_seconds ne '1' ){
					$return_string .= $calc_sub_secs;
				}
				$return_string .= $dt->format_cldr( $last_format_rem ) if $last_format_rem;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"returning: $return_string" ] );
			}
			return $return_string;
		};
	return( $type_filter, $conversion_sub );
}

sub _build_duration{
	my( $self, $duration_ref, $delta_seconds, $delta_nanosecs ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_date::_build_duration', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			'Building a duration string with duration ref:', $duration_ref,
	###LogSD			"With delta seconds: $delta_seconds",
	###LogSD			(($delta_nanosecs) ? "And delta nanoseconds: $delta_nanosecs" : undef) ] );
	my	$return_string;
	my	$key = $duration_ref->[0];
	my	$first = 1;
	for my $position ( 0 .. $duration_ref->[1] ){
		if( $key eq '0' ){
			my $length = length( $last_sub_seconds );
			$return_string .= '.' . sprintf( "%0.${length}f", $delta_nanosecs/1000000000);
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Return string with nanoseconds: $return_string", ] );
		}
		if( $key eq 's' ){
			$return_string .= ( $first ) ? $delta_seconds :
				sprintf "%0$duration_ref->[3]->[$position]d", $delta_seconds;
			$first = 0;
			$key = $duration_order->{$key};
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Delta seconds: $delta_seconds",
			###LogSD		"Next key to process: $key"			] );
		}
		if( $key eq 'm' ){
			my $minutes = int($delta_seconds/60);
			$delta_seconds = $delta_seconds - ($minutes*60);
			$return_string .= ( $first ) ? $minutes :
				sprintf "%0$duration_ref->[3]->[$position]d", $minutes;
			$first = 0;
			$key = $duration_order->{$key};
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Calculated minutes: $minutes",
			###LogSD		"Remaining seconds: $delta_seconds",
			###LogSD		"Next key to process: $key"			] );
		}
		if( $key eq 'h' ){
			my $hours = int($delta_seconds /(60*60));
			$delta_seconds = $delta_seconds - ($hours*60*60);
			$return_string .= ( $first ) ? $hours :
				sprintf "%0$duration_ref->[3]->[$position]d", $hours;
			$first = 0;
			$key = $duration_order->{$key};
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Calculated hours: $hours",
			###LogSD		"Remaining seconds: $delta_seconds",
			###LogSD		"Next key to process: $key"			] );
		}
		$return_string .= $duration_ref->[2]->[$position] if $duration_ref->[2]->[$position];
	}
	return $return_string;
}

sub _build_number{
	my( $self, $type_filter, $list_ref ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Processing a number list to see how it should be converted",
	###LogSD			'With type constraint: ' . $type_filter->name,
	###LogSD			'..using list ref:' , $list_ref 			] );
	my ( $code_hash_ref, $number_type, );
	
	# Resolve zero replacements quickly
	if(	$type_filter->name eq 'ZeroOrUndef' and
		!$list_ref->[-1]->[0] and $list_ref->[-1]->[1] eq '"-"' ){
		###LogSD	$phone->talk( level => 'debug', message =>[
		###LogSD		"Found a zero to bar replacement"			] );
		my $return_string;
		for my $piece ( @$list_ref ){
			$return_string .= $piece->[1];
		}
		$return_string =~ s/"\-"/\-/;
		return( $type_filter, sub{ $return_string } );
	}
	
	# Process once to determine what to do
	for my $piece ( @$list_ref ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"processing number piece:", $piece ] );
		if( defined $piece->[0] ){
			if( my @result = $piece->[0] =~ /^([0-9#\?]+)([,\-\_])?([#0\?]+)?(,+)?$/ ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Regex yielded result:", @result ] );
				my	$comma = ($2) ? $2 : undef,
				my	$comma_less = defined( $3) ? "$1$3" : $1;
				my	$comma_group = length( $3 );
				my	$divide_by_thousands = ( $4 ) ? (( $2 and $2 ne ',' ) ? $4 : "$2$4" ) : undef;#eval{ $2 . $4 }
				my	$divisor = $1 if $1 =~ /^([0-9]+)$/;
				my ( $leading_zeros, $trailinq_zeros );
				if( $comma_less =~ /^[\#\?]*(0+)$/ ){
					$leading_zeros = $1;
				}
				if( $comma_less =~ /^(0+)[\#\?]*$/ ){
					$trailinq_zeros = $1;
				}
				$code_hash_ref->{divide_by_thousands} = length( $divide_by_thousands ) if $divide_by_thousands;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"The comma less string is extracted to: $comma_less",
				###LogSD		((defined $comma_group) ? "The separator group length is: $comma_group" : undef),
				###LogSD		(($comma) ? "The separator character is: $comma" : undef),
				###LogSD		((length( $leading_zeros )) ? ".. w/leading zeros: $leading_zeros" : undef),
				###LogSD		((length( $trailinq_zeros )) ? ".. w/trailing zeros: $trailinq_zeros" : undef),
				###LogSD		(($divisor) ? "..with identified divisor: $divisor" : undef),
				###LogSD		'Initial code hash:', $code_hash_ref] );
				if( !$number_type ){
					$number_type = 'INTEGER';
					$code_hash_ref->{integer}->{leading_zeros} = length( $leading_zeros ) if length( $leading_zeros );
					$code_hash_ref->{integer}->{minimum_length} = length( $comma_less );
					if( $comma ){
						@{$code_hash_ref->{integer}}{ 'group_length', 'comma' } = ( $comma_group, $comma );
					}
					if( defined $piece->[1] ){
						if( $piece->[1] =~ /(\s+)/ ){
							$code_hash_ref->{separator} = $1;
						}elsif( $piece->[1] eq '/' ){
							$number_type = 'FRACTION';
							$code_hash_ref->{numerator}->{leading_zeros} = length( $leading_zeros ) if length( $leading_zeros );
							delete $code_hash_ref->{integer};
						}
					}
				}elsif( ($number_type eq 'INTEGER') or $number_type eq 'DECIMAL' ){
					if( $piece->[1] and $piece->[1] eq '/'){
						$number_type = 'FRACTION';
					}else{
						$number_type = 'DECIMAL';
						$code_hash_ref->{decimal}->{trailing_zeros} = length( $trailinq_zeros ) if length( $trailinq_zeros );
						$code_hash_ref->{decimal}->{max_length} = length( $comma_less );
					}
				}elsif( ($number_type eq 'SCIENTIFIC') or $number_type eq 'FRACTION' ){
					$code_hash_ref->{exponent}->{leading_zeros} = length( $leading_zeros ) if length( $leading_zeros );
					$code_hash_ref->{fraction}->{target_length} = length( $comma_less );
					if( $divisor ){
						$code_hash_ref->{fraction}->{divisor} = $divisor;
					}
				}
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Current number type: $number_type", 'updated settings:', $code_hash_ref] );
			}elsif( $piece->[0] =~ /^((\.)|([Ee][+\-])|(%))$/ ){
				if( $2 ){
					$number_type = 'DECIMAL';
					$code_hash_ref->{separator} = $1;
				}elsif( $3 ){
					$number_type = 'SCIENTIFIC';
					$code_hash_ref->{separator} = $2;
				}else{
					$number_type = 'PERCENT';
				}
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Number type now: $number_type" ] );
			}else{
				confess "badly formed number format passed: $piece->[0]";
			}
		}
	}
	
	my $method = '_build_' . lc( $number_type ) . '_sub';
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Resolved the number type to: $number_type",
	###LogSD		'Working with settings:', $code_hash_ref] );
	my $conversion_sub = $self->$method( $type_filter, $list_ref, $code_hash_ref );
		
	return( $type_filter, $conversion_sub );
}

sub _build_integer_sub{
	my( $self, $type_filter, $list_ref, $code_hash_ref ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_integer_sub', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building an anonymous sub to return integer values",
	###LogSD			'With type constraint: ' . $type_filter->name,
	###LogSD			'..using list ref:' , $list_ref, '..and code hash ref:', $code_hash_ref	] );
	
	my $sprintf_string;
	# Process once to determine what to do
	my $found_integer = 0;
	for my $piece ( @$list_ref ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"processing number piece:", $piece ] );
		if( !$found_integer and defined $piece->[0] ){
			$sprintf_string .= '%s';
			$found_integer = 1;
		}
		if( $piece->[1] ){
			$sprintf_string .= $piece->[1];
		}
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Final sprintf string: $sprintf_string" ] );
	
	my 	$conversion_sub = sub{
			my $adjusted_input = $_[0];
			if( $adjusted_input ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Integer Input: $adjusted_input" ] );
				$adjusted_input = $self->change_output_encoding( $adjusted_input );
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Updated integer Input: $adjusted_input", 'For integer: ' . int( $adjusted_input ) ] );
				if( $type_filter->name eq 'NegativeNum' ){
					$adjusted_input = -1 * $adjusted_input;
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Removed negative: $adjusted_input" ] );
				}
				my $decimal = abs( $adjusted_input - int( $adjusted_input ) );
				$adjusted_input = int( $adjusted_input );
				if( $decimal >= 0.5 ){
					$adjusted_input = ( $adjusted_input < 0 ) ?
						($adjusted_input - 1) : ($adjusted_input + 1);
				}
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Removed the decimal: $decimal", "From adjusted input: $adjusted_input" ] );
				if( exists $code_hash_ref->{integer}->{leading_zeros} and
					length( $adjusted_input ) < $code_hash_ref->{integer}->{leading_zeros} ){
					$adjusted_input = ('0' x ( $code_hash_ref->{integer}->{leading_zeros} - length( $adjusted_input ) ) ) . $adjusted_input;
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Set leading zeros: $adjusted_input" ] );
				}
				if( exists $code_hash_ref->{integer}->{comma} ){
					$adjusted_input = $self->_add_integer_separator( $adjusted_input, @{ $code_hash_ref->{integer}}{ 'comma','group_length' } );
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Added commas: $adjusted_input" ] );
				}
			}else{
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Return undef for empty strings" ] );
				return undef;
			}
			return sprintf( $sprintf_string, $adjusted_input );
		};
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Conversion sub for filter name: " . $type_filter->name, $conversion_sub ] );
	
	return $conversion_sub;
}

sub _build_decimal_sub{
	my( $self, $type_filter, $list_ref, $code_hash_ref ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_decimal_sub', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building an anonymous sub to return decimal values",
	###LogSD			'With type constraint: ' . $type_filter->name,
	###LogSD			'..using list ref:' , $list_ref, '..and code hash ref:', $code_hash_ref	] );
	
	my $sprintf_string;
	# Process once to determine what to do
	for my $piece ( @$list_ref ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"processing number piece:", $piece ] );
		if( defined $piece->[0] ){
			if( $piece->[0] eq '.' ){
				$sprintf_string .= '.';
			}else{
				$sprintf_string .= '%s';
			}
		}
		if( $piece->[1] ){
			$sprintf_string .= $piece->[1];
		}
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Final sprintf string: $sprintf_string" ] );
	
	my 	$conversion_sub = sub{
			my $adjusted_input = $_[0];
			my ( $integer, $decimal );
			if( $adjusted_input ){
				$adjusted_input = $self->change_output_encoding( $adjusted_input );
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Updated integer Input: $adjusted_input" ] );
				if( $type_filter->name eq 'NegativeNum' ){
					$adjusted_input = -1 * $adjusted_input;
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Removed negative: $adjusted_input" ] );
				}
				if( exists $code_hash_ref->{divide_by_thousands} ){
					$adjusted_input = $adjusted_input/( 1000**$code_hash_ref->{divide_by_thousands} );
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Divided by thousands: $adjusted_input" ] );
				}
			}else{
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Return undef for empty strings" ] );
				return undef;
			}
			$integer = int( $adjusted_input );
			$decimal = abs( $adjusted_input - $integer );
			$decimal = undef if !$decimal;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Split input: $adjusted_input", "..to integer: $integer", 
			###LogSD		(($decimal) ? "..and decimal: $decimal" : undef)			 ] );
			if( $decimal ){
				if( exists $code_hash_ref->{decimal}->{max_length} and
					length( $decimal ) > $code_hash_ref->{decimal}->{max_length} ){
					$decimal = sprintf( "%.$code_hash_ref->{decimal}->{max_length}f", $decimal );
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Rounded decimal to: $decimal" ] );
				}
				$decimal =~ /(\d)\.(\d+)/;
				$decimal = $2;
				if( $1 ){
					$integer = ($integer < 0 ) ? ($integer - $1) : ($integer + $1);
				}
				if( exists $code_hash_ref->{decimal}->{trailing_zeros} and
					length( $decimal ) < $code_hash_ref->{decimal}->{trailing_zeros} ){
					$decimal = $decimal . ('0' x ( $code_hash_ref->{decimal}->{trailing_zeros} - length( $decimal ) ) );
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Set trailing zeros: $decimal" ] );
				}
			}else{
				$decimal = '0' x $code_hash_ref->{decimal}->{trailing_zeros};
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Providing default decimal: $decimal" ] );
			}
			if( $integer ){
				if( exists $code_hash_ref->{integer}->{leading_zeros} and
					length( $integer ) < $code_hash_ref->{integer}->{leading_zeros} ){
					$integer = ('0' x ( $code_hash_ref->{integer}->{leading_zeros} - length( $integer ) ) ) . $integer;
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Set leading zeros: $integer" ] );
				}elsif( !$integer ){
					$integer = '0' x $code_hash_ref->{integer}->{leading_zeros};
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Build out integer to leading zeros: $integer" ] );
				}
			}
			if( exists $code_hash_ref->{integer}->{comma} ){
				$integer = $self->_add_integer_separator( $integer, @{ $code_hash_ref->{integer}}{ 'comma', 'group_length' } );
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"Added commas to integer: $integer" ] );
			}
			
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Processing integer: $integer", "..and decimal: $decimal", "..with sprintf: $sprintf_string" ] );
			return sprintf( $sprintf_string, $integer, $decimal);
		};
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Conversion sub for filter name: " . $type_filter->name, $conversion_sub ] );
	
	return $conversion_sub;
}

sub _build_percent_sub{
	my( $self, $type_filter, $list_ref, $code_hash_ref ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_percent_sub', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building an anonymous sub to return decimal values",
	###LogSD			'With type constraint: ' . $type_filter->name,
	###LogSD			'..using list ref:' , $list_ref, '..and code hash ref:', $code_hash_ref	] );
	
	my ( $sprintf_string, $decimal_count );
	# Process once to determine what to do
	for my $piece ( @$list_ref ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"processing number piece:", $piece ] );
		if( defined $piece->[0] ){
			if( $piece->[0] eq '%' ){
				$sprintf_string .= '%%';
			}elsif( $piece->[0] eq '.' ){
				$sprintf_string .= '.';
			}else{
				$sprintf_string .= '%s';
				$decimal_count++;
			}
		}
		if( $piece->[1] ){
			$sprintf_string .= $piece->[1];
		}
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Final sprintf string: $sprintf_string" ] );
	
	my 	$conversion_sub = sub{
			my $adjusted_input = $_[0];
			my ( $integer, $decimal );
			if( $adjusted_input ){
				$adjusted_input = $self->change_output_encoding( $adjusted_input );
				$adjusted_input = $adjusted_input*100;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Updated integer Input: $adjusted_input" ] );
				if( $type_filter->name eq 'NegativeNum' ){
					$adjusted_input = -1 * $adjusted_input;
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Removed negative: $adjusted_input" ] );
				}
			}else{
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Return undef for empty strings" ] );
				return undef;
			}
			$integer = int( $adjusted_input );
			$decimal = abs( $adjusted_input - $integer );
			$decimal = undef if !$decimal;
			if( ($decimal_count == 1) and $decimal and ($decimal >= 0.5) ){
				$integer = ($integer < 0 ) ? ($integer - 1) : ($integer + 1);
			}
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Split input: $adjusted_input", "..to integer: $integer", 
			###LogSD		(($decimal) ? "..and decimal: $decimal" : undef)			 ] );
			if( $integer ){
				if( exists $code_hash_ref->{integer}->{leading_zeros} and
					length( $integer ) < $code_hash_ref->{integer}->{leading_zeros} ){
					$integer = ('0' x ( $code_hash_ref->{integer}->{leading_zeros} - length( $integer ) ) ) . $integer;
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Set leading zeros: $integer" ] );
				}elsif( !$integer ){
					$integer = '0' x $code_hash_ref->{integer}->{leading_zeros};
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Build out integer to leading zeros: $integer" ] );
				}
			}
			if( exists $code_hash_ref->{integer}->{comma} ){
				$integer = $self->_add_integer_separator( $integer, @{ $code_hash_ref->{integer}}{ 'comma', 'group_length' } );
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"Added commas to integer: $integer" ] );
			}
			
			if( $decimal ){
				if( exists $code_hash_ref->{decimal}->{max_length} and
					length( $decimal ) > $code_hash_ref->{decimal}->{max_length} ){
					$decimal = sprintf( "%.$code_hash_ref->{decimal}->{max_length}f", $decimal );
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Rounded decimal to: $decimal" ] );
				}
				$decimal =~ /\.(\d+)/;
				$decimal = $1;
				if( exists $code_hash_ref->{decimal}->{trailing_zeros} and
					length( $decimal ) < $code_hash_ref->{decimal}->{trailing_zeros} ){
					$decimal = $decimal . ('0' x ( $code_hash_ref->{decimal}->{trailing_zeros} - length( $decimal ) ) );
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Set trailing zeros: $decimal" ] );
				}
			}elsif( $decimal_count == 2 ){
				$decimal = '0' x $code_hash_ref->{decimal}->{trailing_zeros};
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Providing default decimal: $decimal" ] );
			}
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Processing integer: $integer",
			###LogSD		(($decimal) ? "..and decimal: $decimal" : undef),
			###LogSD		"..with sprintf: $sprintf_string" ] );
			return sprintf( $sprintf_string, $integer, $decimal);
		};
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Conversion sub for filter name: " . $type_filter->name, $conversion_sub ] );
	
	return $conversion_sub;
}

sub _build_scientific_sub{
	my( $self, $type_filter, $list_ref, $code_hash_ref ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_scientific_sub', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building an anonymous sub to return decimal values",
	###LogSD			'With type constraint: ' . $type_filter->name,
	###LogSD			'..using list ref:' , $list_ref, '..and code hash ref:', $code_hash_ref	] );
	
	my ( $sprintf_string, $exponent_sprintf );
	# Process once to determine what to do
	my	$no_decimal = ( exists $code_hash_ref->{decimal} ) ? 0 : 1 ;
	for my $piece ( @$list_ref ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"processing number piece:", $piece ] );
		if( defined $piece->[0] ){
			if( $piece->[0] =~ /(E)(.)/ ){
				$sprintf_string .= $1;
				$exponent_sprintf = '%';
				$exponent_sprintf .= '+' if $2 eq '+';
				if( exists $code_hash_ref->{exponent}->{leading_zeros} ){
					$exponent_sprintf .= '0.' . $code_hash_ref->{exponent}->{leading_zeros};
				}
				$exponent_sprintf .= 'd';
			}elsif( $piece->[0] eq '.' ){
				$sprintf_string .= '.';
				$no_decimal = 0;
			}elsif( $exponent_sprintf ){
				$sprintf_string .= $exponent_sprintf;
			}else{
				$sprintf_string .= '%s';
			}
		}
		if( $piece->[1] ){
			$sprintf_string .= $piece->[1];
		}
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Final sprintf string: $sprintf_string" ] );
	
	my 	$conversion_sub = sub{
			my $adjusted_input = $_[0];
			if( $adjusted_input ){
				$adjusted_input = $self->change_output_encoding( $adjusted_input );
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Updated scientific Input: $adjusted_input" ] );
				if( $type_filter->name eq 'NegativeNum' ){
					$adjusted_input = -1 * $adjusted_input;
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Removed negative: $adjusted_input" ] );
				}
			}else{
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Return undef for empty strings" ] );
				return undef;
			}
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		"Testing string: " . sprintf( '%0.30f', $adjusted_input ) ] );
			my @results = sprintf( '%0.30f', $adjusted_input ) =~ /^(-)?(\d+)(\.)?(\d+)?/;
			my	$sign 			= $1;
			my	$integer 		= $2;
			my	$place_holder 	= $3;
			my	$decimal		= $4;
			my	$numbers		= $integer . ((defined $decimal) ? $decimal : '');# . '0000000000000000'
			my	$initial_position = length( $integer );
				$numbers =~ /([1-9])/;
			my	$start_real = pos();
			my	$integer_length = $integer % $code_hash_ref->{integer}->{leading_zeros};
				$integer_length ||= $code_hash_ref->{integer}->{leading_zeros};
			$numbers =~ /(0*)(\d{$integer_length})(\d+)?/;
			my	$changed_integer = $1;
				$changed_integer .= $2;
				$integer = $2;
				$decimal = $3;
			my	$adjusted_position = length( $changed_integer );
			my	$exponent = $initial_position - $adjusted_position;
			###LogSD	$phone->talk( level => 'trace', message => [
			###LogSD		"Results of adjusted input are:", @results,
			###LogSD		"Yielding numbers: $numbers", "..initial position: $initial_position",
			###LogSD		(($start_real) ? "..and start of the real numbers: $start_real" : undef),
			###LogSD		"Integer length: $integer_length", "Changed integer: $changed_integer",
			###LogSD		"With new integer: $integer", "New decimal: $decimal", "..and exponent: $exponent" ] );
			my	$test_mod = $exponent % $code_hash_ref->{integer}->{minimum_length};
			if( $test_mod ){
				###LogSD	$phone->talk( level => 'info', message => [
				###LogSD		"The exponent -$exponent- does not divide to the required multiple of: " . $code_hash_ref->{integer}->{minimum_length},
				###LogSD		"There is a remainder of: $test_mod",   ] );
				$exponent -= $test_mod;
				$decimal =~ /(.{$test_mod})(.+)/;
				$integer .= $1;
				$decimal = $2;
				###LogSD	$phone->talk( level => 'info', message => [
				###LogSD		"After adjustement ithe exponent is: $exponent",
				###LogSD		"The integer is: $integer", "The decimal is: $decimal"   ] );
			}
			if( $decimal  ){
				if( exists $code_hash_ref->{decimal}->{max_length} and
					length( $decimal ) > $code_hash_ref->{decimal}->{max_length} ){
					my $test_decimal = '0.' . $decimal;
					$test_decimal = sprintf( "%.$code_hash_ref->{decimal}->{max_length}f", $test_decimal );
					$test_decimal =~ /(\d)\.(\d+)/;
					$integer += $1;
					$decimal = $2;
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Rounded decimal to: $decimal", "With integer: $integer" ] );
				}elsif( $decimal =~ /^[5-9]/ ){
					$integer += 1;
				}
				if( exists $code_hash_ref->{decimal}->{trailing_zeros} and
					length( $decimal ) < $code_hash_ref->{decimal}->{trailing_zeros} ){
					$decimal = $decimal . ('0' x ( $code_hash_ref->{decimal}->{trailing_zeros} - length( $decimal ) ) );
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Set trailing zeros: $decimal" ] );
				}
			}else{
				$decimal = '0' x $code_hash_ref->{decimal}->{trailing_zeros};
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Providing default decimal: $decimal" ] );
			}
			if( $integer ){
				if( exists $code_hash_ref->{integer}->{leading_zeros} and
					length( $integer ) < $code_hash_ref->{integer}->{leading_zeros} ){
					$integer = ('0' x ( $code_hash_ref->{integer}->{leading_zeros} - length( $integer ) ) ) . $integer;
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Set leading zeros: $integer" ] );
				}
				$integer = $sign . $integer if $sign;
			}
			if( exists $code_hash_ref->{integer}->{comma} ){
				$integer = $self->_add_integer_separator( $integer, @{ $code_hash_ref->{integer}}{ 'comma', 'group_length' } );
				###LogSD	$phone->talk( level => 'debug', message =>[
				###LogSD		"Added commas to integer: $integer" ] );
			}
			#~ Should allow no decimal in scientific notiation?  (Excel does!)
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Processing integer: $integer",
			###LogSD		((defined $decimal) ? "..and decimal: $decimal" : undef),
			###LogSD		"and Exponent: $exponent",
			###LogSD		"..with sprintf: $sprintf_string" ] );
			if( $no_decimal ){
				return sprintf( $sprintf_string, $integer, $exponent);
			}else{
				return sprintf( $sprintf_string, $integer, $decimal, $exponent);
			}
		};
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Conversion sub for filter name: " . $type_filter->name, $conversion_sub ] );
	
	return $conversion_sub;
}

sub _build_fraction_sub{
	my( $self, $type_filter, $list_ref, $code_hash_ref ) = @_;
	###LogSD	my	$phone = Log::Shiras::Telephone->new( name_space =>
	###LogSD					$self->get_log_space .  '::_build_number::_build_fraction_sub', );
	###LogSD		$phone->talk( level => 'debug', message => [
	###LogSD			"Building an anonymous sub to return integer and fraction strings",
	###LogSD			'With type constraint: ' . $type_filter->name,
	###LogSD			'..using list ref:' , $list_ref, '..and code hash ref:', $code_hash_ref	] );
	
	my $sprintf_string;
	# Process once to determine what to do
	my $input = 0;
	for my $piece ( @$list_ref ){
		###LogSD	$phone->talk( level => 'debug', message => [
		###LogSD		"processing number piece:", $piece ] );
		if( defined $piece->[0] and $input < 2 ){
			$sprintf_string .= '%s';
			$input++;
		}
		if( $piece->[1] and $piece->[1] ne '/' ){
			$sprintf_string .= $piece->[1];
		}
		###LogSD	$phone->talk( level => 'trace', message => [
		###LogSD		"updated sprintf: $sprintf_string", "Input count: $input" ] );
	}
	my $fract_sprintf = $sprintf_string;
	$fract_sprintf =~ s/%s\s//;
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Final sprintf string: $sprintf_string" ] );
	my $conversion_sub;
	if( exists $code_hash_ref->{fraction}->{divisor} ){
		$conversion_sub = sub{
			my $adjusted_input = $_[0];
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Received conversion request for fixed divisor: " .
			###LogSD		$code_hash_ref->{fraction}->{divisor}, 
			###LogSD		((defined $adjusted_input) ? "..with input: $adjusted_input" : undef) ] );
			my ( $integer, $decimal );
			if( $adjusted_input ){
				$adjusted_input = $self->change_output_encoding( $adjusted_input );
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Updated fraction Input: $adjusted_input" ] );
				if( $type_filter->name eq 'NegativeNum' ){
					$adjusted_input = -1 * $adjusted_input;
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Removed negative: $adjusted_input" ] );
				}
			}else{
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Return undef for empty strings" ] );
				return undef;
			}
			my $sign = '';
			if( $adjusted_input < 0 ){
				$sign = '-';
				$adjusted_input *= -1;
			}
			$integer = int( $adjusted_input );
			$decimal = $adjusted_input - $integer;
			$decimal = undef if !$decimal;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Split input: $adjusted_input", "..to integer: $integer", 
			###LogSD		(($decimal) ? "..and decimal: $decimal" : undef)			 ] );
			
			# Build the fraction
			my $fraction;
			if( $decimal ){
				my $low_numerator = int( $decimal*$code_hash_ref->{fraction}->{divisor} );
				my $high_numerator = $low_numerator + 1;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Calculated low numerator: $low_numerator", 
				###LogSD		"..and high numerator: $high_numerator"			 ] );
				my $numerator = ( 
						($decimal - $low_numerator/$code_hash_ref->{fraction}->{divisor}) < 
						($high_numerator/$code_hash_ref->{fraction}->{divisor} - $decimal)		) ?
							$low_numerator : $high_numerator ;
				if( $numerator == $code_hash_ref->{fraction}->{divisor} ){
					$integer += 1;
					$fraction = undef;
				}elsif( $numerator == 0 ){
					$fraction = undef;
				}else{
					$fraction = $numerator . '/' . $code_hash_ref->{fraction}->{divisor};
				}
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Resolved decimal: $decimal",
				###LogSD		(($fraction) ? "..to fraction: $fraction" : "..to no fraction") ] );
			}else{
				$fraction = undef;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		'No decimal so no fraction' ] );
			}
			if( $integer ){
				if( exists $code_hash_ref->{integer}->{leading_zeros} and
					length( $integer ) < $code_hash_ref->{integer}->{leading_zeros} ){
					$integer = ('0' x ( $code_hash_ref->{integer}->{leading_zeros} - length( $integer ) ) ) . $integer;
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Set leading zeros: $integer" ] );
				}elsif( !$integer ){
					$integer = '0' x $code_hash_ref->{integer}->{leading_zeros};
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Build out integer to leading zeros: $integer" ] );
				}
				if( exists $code_hash_ref->{integer}->{comma} ){
					$integer = $self->_add_integer_separator( $integer, @{ $code_hash_ref->{integer}}{ 'comma', 'group_length' } );
					###LogSD	$phone->talk( level => 'debug', message =>[
					###LogSD		"Added commas to integer: $integer" ] );
				}
				$integer = $sign . $integer;
				if( !$fraction ){
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"No fraction identified - sending just the integer" ] );
					return sprintf( $fract_sprintf, $integer);
				}
			}else{
				if( !$fraction and $adjusted_input ){
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"No resovlvable integer or fraction value - sending zero" ] );
					return sprintf( $fract_sprintf, 0 );
				}
				$fraction = $sign . $fraction;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"No integer found sending just the fraction" ] );
				return sprintf( $fract_sprintf, $fraction);
			}
			
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Processing integer: $integer", "..and fraction: $fraction", "..with sprintf: $sprintf_string" ] );
			return sprintf( $sprintf_string, $integer, $fraction);
		};
	}else{
		$conversion_sub = sub{
			my $adjusted_input = $_[0];
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		'Received conversion request for fixed width fraction' .
			###LogSD		((defined $code_hash_ref->{fraction}->{target_length}) ? (' of width -' . $code_hash_ref->{fraction}->{target_length} . '-') : '') .
			###LogSD		((defined $adjusted_input) ? " with input: $adjusted_input" : '') ] );
			my ( $integer, $decimal );
			if( $adjusted_input ){
				$adjusted_input = $self->change_output_encoding( $adjusted_input );
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Updated fraction Input: $adjusted_input" ] );
				if( $type_filter->name eq 'NegativeNum' ){
					$adjusted_input = -1 * $adjusted_input;
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Removed negative: $adjusted_input" ] );
				}
			}else{
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Return undef for empty strings" ] );
				return undef;
			}
			my $sign = '';
			if( $adjusted_input < 0 ){
				$sign = '-';
				$adjusted_input *= -1;
			}
			$integer = int( $adjusted_input );
			$decimal = $adjusted_input - $integer;
			$decimal = undef if !$decimal;
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Split input: $adjusted_input", "..to integer: $integer", 
			###LogSD		(($decimal) ? "..and decimal: $decimal" : undef)			 ] );
			
			# Build the fraction
			my $fraction;
			if( $decimal ){
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"Processing decimal: $decimal",  ] );
				$fraction = $self->_continuous_fraction( $decimal, 20, $code_hash_ref->{fraction}->{target_length} );
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		(($fraction) ? "Received return fraction: $fraction" : "Decimal too small for a fraction") ] );
				if( defined $fraction ){
					if( $fraction !~ /\// ){
						$integer += $fraction;
						$fraction = undef;
					}
				}else{
					$fraction = undef;
				}
			}else{
				$fraction = undef;
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		'No decimal so no fraction' ] );
			}
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		(($fraction) ? "Finished the decimal with fraction: $fraction" : undef),  ] );
			if( $integer ){
				if( exists $code_hash_ref->{integer}->{leading_zeros} and
					length( $integer ) < $code_hash_ref->{integer}->{leading_zeros} ){
					$integer = ('0' x ( $code_hash_ref->{integer}->{leading_zeros} - length( $integer ) ) ) . $integer;
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Set leading zeros: $integer" ] );
				}elsif( !$integer ){
					$integer = '0' x $code_hash_ref->{integer}->{leading_zeros};
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"Build out integer to leading zeros: $integer" ] );
				}
				if( exists $code_hash_ref->{integer}->{comma} ){
					$integer = $self->_add_integer_separator( $integer, @{ $code_hash_ref->{integer}}{ 'comma', 'group_length' } );
					###LogSD	$phone->talk( level => 'debug', message =>[
					###LogSD		"Added commas to integer: $integer" ] );
				}
				$integer = $sign . $integer;
				if( !$fraction ){
					###LogSD	$phone->talk( level => 'debug', message => [
					###LogSD		"No decimal found sending just the integer" ] );
					return sprintf( $fract_sprintf, $integer);
				}
			}else{
				if( $fraction ){
					$fraction = $sign . $fraction;
				}else{
					$fraction = 0;
				}
				###LogSD	$phone->talk( level => 'debug', message => [
				###LogSD		"No integer found sending just the fraction" ] );
				return sprintf( $fract_sprintf, $fraction);
			}
			
			###LogSD	$phone->talk( level => 'debug', message => [
			###LogSD		"Processing integer: $integer", "..and fraction: $fraction", "..with sprintf: $sprintf_string" ] );
			return sprintf( $sprintf_string, $integer, $fraction);
		};
	}
	###LogSD	$phone->talk( level => 'debug', message => [
	###LogSD		"Conversion sub for filter name: " . $type_filter->name, $conversion_sub ] );
	
	return $conversion_sub;
}

#########1 Phinish            3#########4#########5#########6#########7#########8#########9

no Moose::Role;
	
1;

#########1 Documentation      3#########4#########5#########6#########7#########8#########9
__END__

=head1 NAME

Spreadsheet::XLSX::Reader::LibXML::XMLDOM::Styles - LibXML DOM parser of Styles
    
=head1 DESCRIPTION

This is the module that is used to apply any style definitions listed in the sheet.  See 
L<Spreadsheet::XLSX::Reader::LibXML::Worksheet> for a way to apply other styles to the 
output.  The current styles coverage is minimal and will expand over time.  In general if 
I didn't write the excel version of a style implementation this module will use the 
pass-through style.

=head1 SUPPORT

=over

L<github Spreadsheet-XLSX-Reader-LibXML/issues
|https://github.com/jandrew/Spreadsheet-XLSX-Reader-LibXML/issues>

=back

=head1 TODO

=over

B<1.> Add some L<Data::Walk::Graft> magic to the defined_excel_translations attribute so 
this list can be managed by detail.

B<2.> Add the FmtJapan, FmtJapan2, and FmtUnicode support

There are a lot of features still to be added. This module is very much a work in progress.

=back

=over

=item B<1.> implement more of the L<standard number formats
|http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformat(v=office.14).aspx>

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

B<5.010> - (L<perl>)

L<version>

L<Moose>

L<MooseX::StrictConstructor>

L<MooseX::HasDefaults::RO>

L<XML::LibXML>

L<XML::LibXML::Reader>

L<Type::Coercion>

L<DateTimeX::Format::Excel>

L<Spreadsheet::XLSX::Reader::LogSpace>

L<Spreadsheet::XLSX::Reader::Types>

=back

=head1 SEE ALSO

=over

L<Spreadsheet::XLSX>

L<Spreadsheet::XLSX::Reader::TempFilter>

L<Log::Shiras|https://github.com/jandrew/Log-Shiras>

=back

=cut

#########1#########2 main pod documentation end   5#########6#########7#########8#########9