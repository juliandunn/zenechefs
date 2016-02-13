#!/usr/bin/env ruby

require 'pp'
require 'geocoder'
require 'rubyXL'
require 'ruby_kml'

def comma_or_return(s)
  s.nil? ? '' : (s + ', ')
end

def geocode(l)
  # May need to retry a few times if Google times out
  r = Array.new
  50.downto(1) do
    r = Geocoder.coordinates(l)
    break unless r.nil?
  end
  r
end

workbook = RubyXL::Parser.parse('julian_map_report_2_12_16.xlsx')
worksheet = workbook[0]
worksheet.delete_row(0) # first row is the header
employees = Array.new
worksheet.each do |row|
  if row
    employee = Hash.new
    employee['firstname'] = row[0].value
    employee['lastname'] = row[1].value
    employee['title'] = row[2].value
    employee['department'] = row[3].value
    employee['city'] = row[4].nil? ? nil : row[4].value
    employee['state'] = row[5].nil? ? nil : row[5].value
    employee['zip'] = row[6].nil? ? nil : row[6].value.to_s
    employee['country'] = row[7].value
    employees << employee
  end
end

kml = KMLFile.new
folder = KML::Folder.new(:name => 'Chef Employees')

employees.each do |e|
  location = comma_or_return(e['city']) +
             comma_or_return(e['state']) +
             comma_or_return(e['zip']) +
             e['country']
  e['geocode'] = geocode(location)
  folder.features << KML::Placemark.new(
    :name => e['firstname'] + ' ' + e['lastname'],
    :description => "#{e['title']} - #{e['department']}",
    :geometry => KML::Point.new(:coordinates => {:lat => e['geocode'][0], :lng => e['geocode'][1]})
    )
end
kml.objects << folder
kml.save 'chefemployees.kml'
