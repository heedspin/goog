require 'goog/no_schema_error'
require 'goog/sheet_record_collection'

class Goog::SheetRecord
  class << self
    attr_accessor :schema
  end

  attr :row_values
  attr :key_values
  attr :row_num
  attr :sheet
  attr :schema
  include Goog::SpreadsheetUtils

  def initialize(schema: nil, row_values: nil, row_num: nil, sheet: nil)
    @schema = schema || self.class.schema
    @row_values = row_values
    @row_num = row_num
    @sheet = sheet
    @key_values = {}
  end

  def schema
    @schema || self.class.schema
  end

  def get_row_value(key)
    if @row_values
      if index = self.schema[key]
        @row_values[index]
      else
        nil
      end
    else
      @key_values[key]
    end
  end

  def set_row_value(key, value)
    if @row_values
      if index = self.schema[key]
        @row_values[index] = value
      else
        raise ArgumentError.new("#{key} is not in schema")
      end
    else
      @key_values[key] = value
    end
  end

  def row_values
    raise Goog::NoSchemaError unless self.schema
    if @row_values
      @row_values
    else
      self.schema.to_a.sort_by { |k,i| i }.map { |k,i| @key_values[k] }
    end
  end

  def valid?
    true
  end

  def errors
    @errors ||= []
  end

  def method_missing(mid, *args)
    raise NoMethodError.new("Schema not loaded") if @row_values and self.schema.nil?
    if mid[-1] == '='
      mid = mid[0..-2].to_sym
      raise NoMethodError.new("Unknown method #{mid}=") if @row_values and !self.schema.member?(mid)
      self.set_row_value(mid, args[0])
    else
      mid = mid.to_sym
      raise NoMethodError.new("Unknown method #{mid}") if @row_values and !self.schema.member?(mid)
      self.get_row_value(mid)
    end
  end  

  def self.create_schema(header_row)
    result = {}
    header_row.each_with_index do |value, index|
      result[value.parameterize.underscore.to_sym] = index
    end
    result
  end

  def self.to_collection(values: nil, sheet:, multiple_sheet_values: nil)
    Goog::SheetRecordCollection.new self.from_range_values(values: values, sheet: sheet, multiple_sheet_values: multiple_sheet_values)
  end

  def self.from_range_values(values: nil, sheet:, multiple_sheet_values: nil)
    if values.nil?
      if multiple_sheet_values.nil?
        raise 'values or multiple_sheet_values must be specified'
      end
      values = multiple_sheet_values[sheet.properties.title]
    end
    self.schema = self.create_schema(values[0])
    values[1..-1].map.with_index do |row, index|
      new(schema: self.schema, row_values: row, row_num: index+2, sheet: sheet)
    end
  end

  def self.build(values)
    transaction = new
    values.each do |key, value|
      transaction.send("#{key}=", value)
    end
    transaction
  end

  def self.date_attribute(key)
    self.class_eval <<-RUBY
    def #{key}
      value = self.get_row_value(:#{key})
      return nil unless value.present?
      if value.is_a?(String)
        value = Date.strptime(value, '%m/%d/%Y')
      elsif value.is_a?(Integer)
        value = Date.new(1899,12,30).advance(days: value)
      end
      value
    end

    def #{key}=(value)
      if value.is_a?(Date)
        value = value.strftime('%m/%d/%Y') 
      end
      self.set_row_value :#{key}, value
    end
    RUBY
  end

  def self.active_hash_attribute(key, klass)
    if !klass.is_a?(String)
      klass = klass.name
    end
    self.class_eval <<-RUBY
    def #{key}
      #{klass}.find_by_name self.get_row_value(:#{key})
    end

    def #{key}=(value)
      value = #{klass}.to_object(value)
      self.set_row_value :#{key}, value.name
    end
    RUBY
  end

  def self.float_attribute(key)
    self.class_eval <<-RUBY
    def #{key}
      value = self.get_row_value(:#{key})
      return nil unless value.present?
      if value.is_a?(String)
        value = value.gsub(/[^\d^\.]/, '').to_f
      end
      value
    end
    RUBY
  end
end