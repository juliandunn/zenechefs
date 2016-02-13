#!/usr/bin/env ruby

require 'pp'
require 'rexml/document'
require 'geocoder'
require 'ruby_kml'

doc = REXML::Document.new File.new('Zenefits.xml')
i = 0
employees = Array.new
rec = Array.new
geocodes = Hash.new

def geocode_or_return(geocodes, dest)
  return geocodes[dest] unless geocodes[dest].nil?
  # May need to retry a few times if Google times out
  20.downto(1) do
    geocodes[dest] = Geocoder.coordinates(dest)
    unless geocodes[dest].nil?
      break
    end
  end
  return geocodes[dest]
end

REXML::XPath.each(doc, "//tbody[@class='ember-view']/tr/td") do |e|
  if e.has_elements? # It's the block with the username or name
    REXML::XPath.each(e, "div[@class='primary']/a") do |f|
      rec << f.text.gsub(Regexp.new(/[\s]+/), ' ')
    end
    REXML::XPath.each(e, "div[@class='ember-view']/a") do |f|
      rec << f.text
    end
    REXML::XPath.each(e, "div[@class='ember-view']") do |f|
      rec << f.text.gsub(Regexp.new(/[\s]+/), '')
    end
  else
    rec << e.text.gsub('Remote - ', '').gsub('Seattle', 'Seattle, WA')
  end
  i += 1
  if i == 6
    rec.reject! { |r| r.empty? || r == 'N/A' }  # clean up crap
    rec << geocode_or_return(geocodes, rec[3])
    employees << rec
    rec = []
    i = 0
  end
end

kml = KMLFile.new

folder = KML::Folder.new(:name => 'Chef Employees')
employees.each do |employee|
  folder.features << KML::Placemark.new(
    :name => employee[0],
    :description => "#{employee[1]} - #{employee[2]}",
    :geometry => KML::Point.new(:coordinates => {:lat => employee[-1][0], :lng => employee[-1][1]})
    )
end
kml.objects << folder
kml.save 'chefemployees.kml'
