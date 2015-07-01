#!/usr/bin/ruby

require 'csv'
require 'openssl'
require 'optparse'
require 'net/http'
require 'uri'

TWEETID_COL = 0
TWEETTEXT_COL = 2

@options = {}
OptionParser.new do |opts|
  opts.banner = 'Usage: twitLinks.rb [options] file1 file2 ...'
  @options[:append] = nil
  @options[:encoding] = 'ISO8859-1'
  @options[:extension] = '.csv'
  @options[:redirects] = 5

  opts.on( '-a', '--append FILE', 'Append output in a single file, which may exist' ) {|a| @options[:append] = a}
  opts.on( '-e', '--encoding ENC', "Open source files using a specific encoding (default: #{@options[:encoding]})" ) {|e| @options[:encoding] = e}
  opts.on( '-h', '--help', 'Display this usage information' ) {puts opts; exit}
  opts.on( '-m', '--mentions', 'Generate separate output file(s) tracking mentions' ) {|m| @options[:mentions] = m}
  opts.on( '-r', '--redirects MAX', "Specify maximum redirects allowed per link (default: #{@options[:redirects]})" ) {|r| @options[:redirects] = r}
  opts.on( '-t', '--tags', 'Generate separate output file(s) tracking hashtags' ) {|t| @options[:tags] = t}
  opts.on( '-v', '--[no-]verbose', 'Display extra info during execution' ) {|v| @options[:verbose] = v}
  opts.on( '-x', '--extension EXT', "Use a specific extension for output files (default: #{@options[:extension]})" ) {|x| @options[:extension] = x}
end.parse!

# Construct an output path from the filename given. This amounts to adding
# '_out' (or other descriptor) to the root filename and changing the exten-
# sion, if necessary.
def get_output_path( file, defExt, type = 'out' )
  arr = File.split( file )
  ext = File.extname( arr[-1] )

  arr[-1] = "#{File.basename( arr[-1], ext )}_#{type}#{defExt}"
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


# Add an extension to the target file name, if it doesn't have one already.
if @options[:append]
  @options[:append] += @options[:extension] if File.extname( @options[:append] ).empty?
end

# Display some descriptive text about this invocation, if appropriate.
if @options[:verbose]
  puts 'Being verbose'
  puts "Using #{@options[:encoding]} encoding"
  puts "Default extension is #{@options[:extension]}"
  puts "Maximum of #{@options[:redirects]} redirects allowed per link"
  puts "Appending files to #{@options[:append]}" if @options[:append]
end

def process
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
