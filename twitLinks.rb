#!/usr/bin/ruby

require 'csv'
require 'openssl'
require 'optparse'
require 'net/http'
require 'uri'

TWEETTEXT_COL = 2

options = {}
OptionParser.new do |opts|
  opts.banner = 'Usage: twitLinks.rb [options] file1 file2 ...'
  options[:append] = nil
  options[:encoding] = 'ISO8859-1'
  options[:extension] = '.csv'

  opts.on( '-a', '--append FILE', 'Append output in a single file, which may exist' ) {|a| options[:append] = a}
  opts.on( '-e', '--encoding ENC', "Open source files using a specific encoding (default: #{options[:encoding]})" ) {|e| options[:encoding] = e}
  opts.on( '-h', '--help', 'Display this usage information' ) {puts opts; exit}
  opts.on( '-v', '--[no-]verbose', 'Display extra info during execution' ) {|v| options[:verbose] = v}
  opts.on( '-x', '--extension EXT', "Use a specific extension for output files (default: #{options[:extension]})" ) {|x| options[:extension] = x}
end.parse!

# Follow an unlimited number of redirect requests to eventually reach a
# final resource location on the web.
def redirect( uri )
  location = nil

  until uri.nil?
    location = uri
    uri = get_location( location )
  end

  location 
end

# Do a HEAD request for the specified URI and look for a redirect location
# header in the response.
def get_location( uri )
  uri = URI.parse( uri )
  http = Net::HTTP.new( uri.host, uri.port )

  # If the target uses SSL, make sure our request does, too.
  if 'https' == uri.scheme
    http.use_ssl = true
    http.verify_mode = OpenSSL::SSL::VERIFY_NONE
  end

  # Perform the request and retrieve the location header.
  req = Net::HTTP::Head.new( uri.request_uri )
  loc = http.request( req )['location']

  # If the location header is present, make sure it is a fully formed URI.
  rebase( loc, uri ) if loc
rescue => e
  puts e.to_s
  nil
end

# Supply URI components from the original source if a redirect location is
# missing them.
def rebase( location, uri )
  target = URI.parse( URI.encode( location ) )
  target.scheme ||= uri.scheme
  target.host ||= uri.host
  target.port ||= uri.port
  target.to_s
end

# Construct an output path from the filename given. This amounts to adding
# '_out' to the root filename and changing the extension, if necessary.
def get_output_path( file, defExt )
  arr = File.split( file )
  ext = File.extname( arr[-1] )

  arr[-1] = "#{File.basename( arr[-1], ext )}_out#{defExt}"
  File.join( arr )
end


# Add an extension to the target file name, if it doesn't have one already.
if options[:append]
  options[:append] += options[:extension] if File.extname( options[:append] ).empty?
end

# Display some descriptive text about this invocation, if appropriate.
if options[:verbose]
  puts 'Being verbose'
  puts "Using #{options[:encoding]} encoding"
  puts "Default extension is #{options[:extension]}"
  puts "Appending files to #{options[:append]}" if options[:append]
end

# Process each file specified on the command line.
ARGV.each do |file|
  file += options[:extension] if File.extname( file ).empty?
  puts "Reading #{file}..." if options[:verbose]

  links = CSV.read( file, encoding: options[:encoding] )
  out = options[:append] || get_output_path( file, options[:extension] )

  CSV.open( out, 'ab' ) do |csv|
    puts "Writing to #{out}..." if options[:verbose]

    links.each_with_index do |row, i|
      row << 'URL' if 0 == i

      # Eliminate carriage returns in the source data, since they contaminate
      # the output and lead to formatting errors.
      row.map! {|col| col.gsub( /\r/, '' )}
      target = nil
  
      # Find the first link embedded in the tweet text, if any.
      if /^.*?(?<url>https*:\/\/t\.co\/[a-zA-Z0-9]+)/.match( row[TWEETTEXT_COL] )
        # Follow all redirects until the actual resource is found.
        url = $~[:url] 
        uri = URI.parse( redirect( url ) )
        target = uri.to_s

        # If the resulting target points back to Twitter, it's an embedded
        # media resource.
        if /^https*:\/\/twitter\.com\/.+\/[0-9]+\/(?<media>.+)\//.match( target )
          target = "<embedded #{$~[:media]} media>"
        end

        puts "  #{i}:#{url} --> #{target}" if options[:verbose]
      else
        puts "  #{i}:<no links>" if options[:verbose]
      end

      row << target
      csv << row
    end
  end
end
