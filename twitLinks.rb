#!/usr/bin/ruby

require 'csv'
require 'openssl'
require 'optparse'
require 'net/http'
require 'rubyXL'
require 'uri'


SHEET_NAMES   = {twitter:'From Twitter', links:'Links', mentions:'Mentions', hashtags:'Hashtags'}
TWITTER_COLS  = ['Tweet id', 'Tweet permalink', 'Tweet text', 'time',
                 'impressions', 'engagements', 'engagement rate', 'retweets',
                 'replies', 'favorites', 'user profile clicks', 'url clicks',
                 'hashtag clicks', 'detail expands', 'permalink clicks',
                 'embedded media views', 'app opens', 'app installs', 'follows']
LINKS_COLS    = ['Tweet id', 'link', 'original']
MENTIONS_COLS = ['Tweet id', 'mention']
HASHTAGS_COLS = ['Tweet id', 'hashtag']
PROGRESS_LEN  = 25
  
@options = {}
OptionParser.new do |opts|
  opts.banner = 'Usage: twitLinks.rb [options] file1 file2 ...'
  @options[:encoding] = 'UTF-8'
  @options[:output] = nil
  @options[:redirects] = 5

  opts.on( '-e', '--encoding ENC', "Open source files using a specific encoding (default: #{@options[:encoding]})" ) {|e| @options[:encoding] = e}
  opts.on( '-h', '--help', 'Display this usage information' ) {puts opts; exit}
  opts.on( '-o', '--output FILE', 'Combine results in a single output file' ) {|o| @options[:output] = o}
  opts.on( '-r', '--redirects MAX', "Specify maximum redirects allowed per link (default: #{@options[:redirects]})" ) {|r| @options[:redirects] = r.to_i}
  opts.on( '-t', '--[no-]truncate', 'Overwrite output file if they exist' ) {|t| @options[:truncate] = t}
  opts.on( '-v', '--[no-]verbose', 'Display extra info during execution' ) {|v| @options[:verbose] = v}
end.parse!

def add_ext( name, ext )
  arr = File.split( name )
  arr[-1] = File.basename( arr[-1], File.extname( arr[-1] ) ) + '.' + ext
  File.join( arr )
end

# Trace a URI to its final location, following redirects as necessary (but
# only up to some maximum number of hops).
def find_target( uri, maxRedirects = @options[:redirects] )
  if 0 < maxRedirects
    http = Net::HTTP.new( uri.host, uri.port )

    # If the target uses SSL, make sure our request does, too.
    if 'https' == uri.scheme
      http.use_ssl = true
      http.verify_mode = OpenSSL::SSL::VERIFY_NONE
    end

    # Perform the request and retrieve the location header.
    req = Net::HTTP::Head.new( uri.request_uri )
    loc = http.request( req )['location']

    # If there is a location header, make sure the target uri comprises all
    # components necessary to redirect.
    if loc
      redirect = URI.parse( URI.encode( loc ) )
      if URI::Generic == redirect.class
        redirect.scheme ||= uri.scheme
        redirect.host ||= uri.host
        redirect.port ||= uri.port
        redirect = URI.parse( redirect.to_s )
      end

      # Recursively try again.
      uri = find_target( redirect, maxRedirects - 1 )
    end
  else
    @tooDeep += 1
  end

  uri
rescue => e
  puts e.to_s
  nil
end

def get_workbook( filename )
  workbook = nil

  # Try to open an existing workbook if one is specified and we haven't been
  # asked to simply overwrite it.
  begin
    workbook = RubyXL::Parser.parse( filename ) unless @options[:truncate]
  rescue
    puts "  Output file #{filename} missing or invalid... creating" if @options[:verbose]
  end

  # Create a new workbook if necessary, then add tabs for the data we will
  # output, if they don't exist.
  if workbook.nil?
    workbook = RubyXL::Workbook.new
    workbook.worksheets.shift
  end
  SHEET_NAMES.values.each {|tab| workbook.add_worksheet( tab ) if workbook[tab].nil?}

  workbook
end

# Copy over the original Twitter data that we care about.
def write_twitter_row( workbook, tweet, indices )
  id = tweet[1].scan( /\d+$/ )[0]
  text = tweet[2].gsub( /\r/, '' ) 

  worksheet = workbook[SHEET_NAMES[:twitter]]
  row = worksheet.count

  # Add the header row if it doesn't already exist. Usually not an issue if
  # updating an existing spreadsheet. 
  if 0 == row
    TWITTER_COLS.each_with_index {|label, i| worksheet.add_cell( row, i, label )}
    row += 1
  end

  # Substitute our parsed id for the one Twitter provides (since it's worth-
  # less) and the sanitized tweet text, then copy the rest of the data as-is.
  indices.each_with_index {|i, j| worksheet.add_cell( row, j, tweet[i] )}
  worksheet[row][0].change_contents( id )
  worksheet[row][2].change_contents( text )

  return id, text
end

def write_link_rows( workbook, id, text )
  worksheet = workbook[SHEET_NAMES[:links]]
  row = worksheet.count

  # Add the header row if it doesn't already exist. Usually not an issue if
  # updating an existing spreadsheet. 
  if 0 == row
    LINKS_COLS.each_with_index {|label, i| worksheet.add_cell( row, i, label )}
    row += 1
  end

  # Output a row mapping each tweet id to the links it contains.
  text.scan( /https*:\/\/t\.co\/\w+/ ).each do |link|
    worksheet.add_cell( row, 0, id )
    worksheet.add_cell( row, 1, find_target( URI.parse( link ) ) )
    worksheet.add_cell( row, 2, link )
    row += 1
  end
end

def write_mention_rows( workbook, id, text )
  worksheet = workbook[SHEET_NAMES[:mentions]]
  row = worksheet.count

  # Add the header row if it doesn't already exist. Usually not an issue if
  # updating an existing spreadsheet.
  if 0 == row
    MENTIONS_COLS.each_with_index {|label, i| worksheet.add_cell( row, i, label )}
    row += 1
  end

  # Output a row mapping each tweet id to the mentions it contains.
  text.scan( /@(\w{1,15})/ ).flatten.each do |mention|
    worksheet.add_cell( row, 0, id )
    worksheet.add_cell( row, 1, mention )
    row += 1
  end
end

def write_hashtag_rows( workbook, id, text )
  worksheet = workbook[SHEET_NAMES[:hashtags]]
  row = worksheet.count

  # Add the header row if it doesn't already exist. Usually not an issue if
  # updating an existing spreadsheet. 
  if 0 == row
    HASHTAGS_COLS.each_with_index {|label, i| worksheet.add_cell( row, i, label )}
    row += 1
  end

  # Output a row mapping each tweet id to the hashtags it contains.
  text.scan( /#(\w+)/ ).flatten.each do |hashtag|
    worksheet.add_cell( row, 0, id )
    worksheet.add_cell( row, 1, hashtag )
    row += 1
  end
end

def read_csv( file )
  filename = add_ext( file, 'csv' )
  puts "Reading #{filename}..."

  CSV.read( filename, encoding: @options[:encoding] )
end

def write_xlsx( file, tweets )
  filename = add_ext( file, 'xlsx' )
  puts "Writing #{filename}..."

  workbook = get_workbook( filename )
  indices = TWITTER_COLS.map {|name| tweets[0].find_index( name )}
  tweets.shift

  puts "  Processing #{tweets.length} tweets:" if @options[:verbose]
  @tooDeep = 0

  tweets.each_with_index do |tweet, i|
    id, text = write_twitter_row( workbook, tweet, indices )

    pct = 100.0 * i / tweets.length
    prog = (PROGRESS_LEN * pct / 100.0).to_i
    print "  [%s>%s] %d%% (#{id})    \r" % ['=' * prog, ' ' * (PROGRESS_LEN - prog), pct] if @options[:verbose] 

    write_link_rows( workbook, id, text )
    write_mention_rows( workbook, id, text )
    write_hashtag_rows( workbook, id, text )
  end

  if @options[:verbose]
    puts "  [%s>] 100%%%s" % ['=' * PROGRESS_LEN, ' ' * PROGRESS_LEN]
    puts "  #{@tooDeep} link#{'s' unless 1 == @tooDeep} exceeded redirect limit (max #{@options[:redirects]})" if 0 < @tooDeep
  end

  workbook.write( filename )
rescue => e
  puts e.to_s
end


# Display some descriptive text about this invocation, if appropriate.
if @options[:verbose]
  puts 'Being verbose'
  puts "Using #{@options[:encoding]} encoding"
  puts "Maximum of #{@options[:redirects]} redirects allowed per link"

  puts "Combining output in a single file" if @options[:output]
  puts "Overwriting output file#{'s' if @options[:output]}" if @options[:truncate]
end

if @options[:output]
  tweets = []
  ARGV.each {|file| tweets += read_csv( file )}
  write_xlsx( @options[:output], tweets )
else
  ARGV.each {|file| write_xlsx( File.basename( file ), read_csv( file ) )}
end
