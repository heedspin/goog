require 'goog/no_schema_error'
require 'goog/sheet_record_collection'

class Goog::SheetRecord
  include Goog::ColumnToLetter
  class << self
    attr_accessor :schema
  end

  attr :row_values
  attr :key_values
  attr :row_num
  attr :spreadsheet_id
  attr :sheet
  attr :schema
  attr :changes
  attr :major_dimension

  def initialize(schema: nil, row_values: nil, row_num: nil, sheet: nil, spreadsheet_id: nil, major_dimension: nil)
    @schema = schema || self.class.schema
    @row_values = row_values
    @row_num = row_num
    @spreadsheet_id = spreadsheet_id
    @sheet = sheet
    @major_dimension = major_dimension
    @key_values = {}
  end

  def schema
    @schema || self.class.schema
  end

  def a1(column)
    index = column.is_a?(Symbol) ? self.schema[column] : column
    raise "Unknown column #{column}" if index.nil?
    result = []
    if @sheet
      result.push (@sheet.is_a?(String) ? @sheet : @sheet.properties.title) + '!'
    end
    if @major_dimension.nil? or (@major_dimension == :rows)
      result.push Goog::Services.sheets.column_to_letter(index+1)
      result.push self.row_num.to_s
    else
      result.push self.column_to_letter(self.row_num)
      result.push (index + 1).to_s
    end
    result.join
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
    previous_value = self.get_row_value(key)
    if previous_value != value
      (@changes ||= []).push [key, previous_value, value]
    end
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

  def changes_to_s
    if @changes.nil?
      ''
    else
      @changes.map { |key, previous_value, new_value|
        "#{key} #{previous_value} => #{new_value}"
      }.join(', ')
    end
  end
  def changed?
    @changes.present?
  end

  def self.ensure_schema(spreadsheet_id: nil, sheet: nil, major_dimension: :rows)
    if !self.schema
      if spreadsheet_id and sheet and major_dimension
        self.load_schema(spreadsheet_id, sheet: sheet, major_dimension: major_dimension)
      end
      raise Goog::NoSchemaError unless self.schema
    end
    self.schema
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

  def save
    Goog::Services.sheets.write_changes(self.spreadsheet_id, self, major_dimension: self.major_dimension)
  end

  def self.create_schema(header_row)
    result = {}
    header_row.each_with_index do |value, index|
      key = value.parameterize.underscore.to_sym
      if @@rename_columns and (new_key = @@rename_columns[key])
        key = new_key
      end
      result[key] = index
    end
    result
  end

  def self.load_schema(spreadsheet_id, sheet: nil, value_render_option: :unformatted_value, major_dimension: :rows)
    range = if major_dimension == :rows
      '1:1'
    else
      'A:A'
    end
    values = Goog::Services.sheets.get_range(spreadsheet_id, range, sheet: sheet, value_render_option: value_render_option, major_dimension: major_dimension)
    self.from_range_values(values: values, spreadsheet_id: spreadsheet_id, sheet: sheet, major_dimension: major_dimension)
  end

  def self.find(spreadsheet_id, range: nil, sheet: nil, value_render_option: :unformatted_value, major_dimension: :rows)
    if sheet
      sheet = Goog::Services.sheets.get_sheet_by_name(spreadsheet_id, sheet) if sheet.is_a?(String)
    else
      sheet = Goog::Services.sheets.get_sheets(spreadsheet_id).first
    end
    values = Goog::Services.sheets.get_range(spreadsheet_id, range, sheet: sheet, major_dimension: major_dimension)
    self.from_range_values(values: values, spreadsheet_id: spreadsheet_id, sheet: sheet, major_dimension: major_dimension)
  end

  def self.find_all(spreadsheet_id, sheet: nil, value_render_option: :unformatted_value, major_dimension: :rows)
    if sheet
      sheet = Goog::Services.sheets.get_sheet_by_name(spreadsheet_id, sheet)
    end
    values = Goog::Services.sheets.get_range(spreadsheet_id, sheet: sheet, value_render_option: value_render_option, major_dimension: major_dimension)
    self.from_range_values(values: values, spreadsheet_id: spreadsheet_id, sheet: sheet, major_dimension: major_dimension)
  end

  def self.to_collection(values: nil, spreadsheet_id: nil, sheet: nil, multiple_sheet_values: nil)
    Goog::SheetRecordCollection.new self.from_range_values(values: values, spreadsheet_id: spreadsheet_id, sheet: sheet, multiple_sheet_values: multiple_sheet_values)
  end

  def self.from_range_values(values: nil, spreadsheet_id: nil, sheet: nil, multiple_sheet_values: nil, major_dimension: :rows)
    if values.nil?
      if multiple_sheet_values.nil?
        raise 'values or multiple_sheet_values must be specified'
      end
      sheet_title = sheet.is_a?(String) ? sheet : sheet.properties.title
      values = multiple_sheet_values[sheet_title]
    end
    self.schema = self.create_schema(values[0])
    values[1..-1].map.with_index do |row, index|
      new(schema: self.schema, 
          row_values: row, 
          row_num: index+2, 
          spreadsheet_id: spreadsheet_id, 
          sheet: sheet,
          major_dimension: major_dimension)
    end
  end

  def write_attribute(key, value)
    self.send("#{key}=", value)
  end
  def set_attributes(values)
    values.each do |key, value|
      self.write_attribute(key, value)
    end
    self
  end

  def respond_to?(mid)
    self.schema.member?(mid)
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

  @@rename_columns = nil
  def self.rename_columns(map)
    @@rename_columns = map
  end

  def self.build(values)
    new.set_attributes(values)
  end

  GOOGLE_EPOCH=Date.new(1899,12,30)

  def self.date_attribute(key)
    self.class_eval <<-RUBY
    def #{key}
      value = self.get_row_value(:#{key})
      return nil unless value.present?
      if value.is_a?(String)
        value = Date.strptime(value, '%m/%d/%Y')
      elsif value.is_a?(Integer)
        value = GOOGLE_EPOCH.advance(days: value)
      end
      value
    end

    def #{key}=(value)
      if value.is_a?(Date)
        # value = value.strftime('%m/%d/%Y') 
        value = (value - GOOGLE_EPOCH).to_i
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
      #{klass}.find_by_alias self.get_row_value(:#{key})
    end

    def #{key}=(value)
      value = #{klass}.to_object(value)
      self.set_row_value :#{key}, value.try(:name)
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