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
  
@options = {}
OptionParser.new do |opts|
  opts.banner = 'Usage: twitLinks.rb [options] file1 file2 ...'
  @options[:encoding] = 'UTF-8'
  @options[:output] = nil
  @options[:redirects] = 5

  opts.on( '-e', '--encoding ENC', "Open source files using a specific encoding (default: #{@options[:encoding]})" ) {|e| @options[:encoding] = e}
  opts.on( '-h', '--help', 'Display this usage information' ) {puts opts; exit}
  opts.on( '-o', '--output FILE', 'Combine results in a single output file' ) {|o| @options[:output] = o}
  opts.on( '-r', '--redirects MAX', "Specify maximum redirects allowed per link (default: #{@options[:redirects]})" ) {|r| @options[:redirects] = r}
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
      redirect.scheme ||= uri.scheme
      redirect.host ||= uri.host
      redirect.port ||= uri.port

      # Recursively try again.
      uri = find_target( redirect, maxRedirects - 1 )
    end
  else
    puts "Redirect limit exceeded (maximum=#{@options[:redirects]})"
  end

  uri
rescue => e
  puts e.to_s
  nil
end

# Extract all the links, mentions, and hashtags from some tweet text.
def parse_tweet( text )
  links = text.scan( /https*:\/\/t\.co\/\w+/ ).map {|link| find_target( URI.parse( link ) )}
  mentions = text.scan( /@(\w{1,15})/ ).flatten
  tags = text.scan( /#(\w+)/ ).flatten
  
  return links, mentions, tags
end

def get_workbook( filename )
  workbook = nil

  # Try to open an existing workbook if one is specified and we haven't been
  # asked to simply overwrite it.
  begin
    workbook = RubyXL::Parser.parse( filename ) unless @options[:truncate]
  rescue
    puts "  Output file #{file} missing or invalid... creating" if @options[:verbose]
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

def read_csv( file )
  filename = add_ext( file, 'csv' )
  puts "Reading #{filename}..."

  CSV.read( filename, encoding: @options[:encoding] )
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

  # Output a row mapping each tweet id to the mentions it contains.
  text.scan( /#(\w+)/ ).flatten.each do |hashtag|
    worksheet.add_cell( row, 0, id )
    worksheet.add_cell( row, 1, hashtag )
    row += 1
  end
end

def write_xlsx( file, tweets )
  filename = add_ext( file, 'xlsx' )
  puts "Writing #{filename}..."

  workbook = get_workbook( filename )
  indices = TWITTER_COLS.map {|name| tweets[0].find_index( name )}

  tweets.shift
  tweets.each do |tweet|
    id, text = write_twitter_row( workbook, tweet, indices )
    write_link_rows( workbook, id, text )
    write_mention_rows( workbook, id, text )
    write_hashtag_rows( workbook, id, text )
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

  puts "Writing combining output to #{@options[:output]}" if @options[:output]
  puts "Overwriting output file#{'s' if @options[:output]}" if @options[:truncate]
end

def process
  if @options[:output]
    tweets = []
    ARGV.each {|file| tweets += read_csv( file )}
    write_xslx( @options[:output], tweets )
  else
    ARGV.each {|file| write_xslx( file, read_csv( file ) )}
  end
end


def proc2
# Process each file specified on the command line.
ARGV.each do |file|
  mentions, tags = [], []
  file += @options[:extension] if File.extname( file ).empty?
  puts "Reading #{file}..."

  links = CSV.read( file, encoding: @options[:encoding] )
  out = @options[:append] || get_output_path( file, @options[:extension] )

  CSV.open( out, 'ab' ) do |csv|
    puts "Writing to #{out}..."

    links.each_with_index do |row, i|
      row << 'URL' if 0 == i
      id = row[TWEETID_COL]

      # Eliminate carriage returns in the source data, since they contaminate
      # the output and lead to formatting errors.
      row.map! {|col| col.gsub( /\r/, '' )}
      tweet = row[TWEETTEXT_COL]
      target = nil
  
      # Find the first link embedded in the tweet text, if any.
      if /^.*?(?<url>https*:\/\/t\.co\/[a-zA-Z0-9]+)/.match( tweet )
        # Follow all redirects until the actual resource is found.
        url = $~[:url] 
        uri = URI.parse( redirect( url ) )
        target = uri.to_s

        # If the resulting target points back to Twitter, it's an embedded
        # media resource.
        if /^https*:\/\/twitter\.com\/.+\/[0-9]+\/(?<media>.+)\//.match( target )
          target = "<embedded #{$~[:media]} media>"
        end

        puts "  #{i}:#{url} --> #{target}" if @options[:verbose]
      else
        puts "  #{i}:<no links>" if @options[:verbose]
      end

      # Track hashtags and references to other Twitter users, if requested.
      tweet.scan( /@(\w{1,15})/ ).flatten.each {|user| mentions << [id, user]} if @options[:mentions]
      tweet.scan( /#(\w+)/ ).flatten.each {|tag| tags << [id, tag]} if @options[:tags]
 
      row << target
      csv << row
    end
  end

  # Write a separate output file tracking mentions, if requested.
  if @options[:mentions]
    out = get_output_path( file, @options[:extension], 'mentions' )
    CSV.open( out, 'ab' ) do |csv|
      puts "Writing mentions to #{out}..."
      csv << ['Tweet id', 'Mention']
      mentions.each {|mention| csv << mention}
    end
  end

  # Write a separate output file tracking hashtags, if requested.
  if @options[:tags]
    out = get_output_path( file, @options[:extension], 'hashtags' )
    CSV.open( out, 'ab' ) do |csv|
      puts "Writing hashtags to #{out}..."
      csv << ['Tweet id', 'Hashtag']
      tags.each {|tag| csv << tag}
    end
  end
end
end
