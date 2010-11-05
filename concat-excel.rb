#!/usr/bin/env ruby
require 'rubygems'
require 'spreadsheet'
require 'pp'

STDOUT.sync = true

class ConcatApp
  def self.run
    xl_joiner = XLJoiner.new ARGV.pop
    xl_joiner.join ARGV
  end
end

class XLJoiner
  def initialize destination
    @destination = destination
    @workbook = Spreadsheet::Workbook.new
    @target_worksheet = @workbook.create_worksheet
    @current_worksheet_columns = {}
    @column_list = ['file', 'worksheet']
    @data = []
  end

  def join files
    files.each do |file|
      @file = file
      $stdout.write "Processing #{file} ... "
      current_book = Spreadsheet.open file
      current_book.worksheets.each do |current_worksheet|
        @worksheet = current_worksheet.name
        catch (:no_content) do
          index_columns_in current_worksheet
          slurp_data_from current_worksheet
        end
      end
      $stdout.write "[DONE]\n"
    end
    $stdout.write "Writing out ..."
    write_out
    $stdout.write "[DONE]"
  end

  def find_header_of worksheet
    index = 0
    puts "-"
    worksheet.each 0 do |row|
      headers = row.reject { |col| col.to_s.strip.length == 0 }
      # $stdout.puts "#{@worksheet}[#{index}]: #{row.join(' | ')}" if headers.length > 5
      return index if headers.length > 5
      index = index + 1
    end
    throw :no_content
  end

  def index_columns_in worksheet
    @current_worksheet_columns = {}
    @header_index = find_header_of worksheet
    header_row = worksheet.row @header_index
    index = 0
    puts "#{@worksheet}[#{@header_index}]: | #{header_row.join(' | ')} |"
    header_row.each do |col_name|
      column_name = col_name.to_s.strip
      @current_worksheet_columns[column_name] = index
      @column_list.push column_name unless @column_list.include? column_name
      index = index + 1
    end
  end

  def slurp_data_from worksheet
    worksheet.each @header_index + 1 do |row|
      target_row = [@file, @worksheet]
      @current_worksheet_columns.each_pair do |column_name,column_index|
        target_row[@column_list.index(column_name)] = row[column_index]
      end
      @data.push target_row unless invalid_row? target_row
    end
  end

  def write_out
    @data.unshift @column_list
    puts "#{@data.length} rows"
    File.open('data.yaml', 'w') { |out| YAML.dump(@data,out) }
    @data.each_index do |row_index|
      @data[row_index].each_index do |column_index|
        @target_worksheet[row_index,column_index] = @data[row_index][column_index] unless @data[row_index][column_index].nil?
      end unless @data[row_index].nil?
    end
    @workbook.write @destination
  end

  def invalid_row? row
    row.each_index do |index|
      next if index == 0 or index == 1
      return false if not row[index].nil?
    end
    return true
  end
  
end

ConcatApp.run


