require 'goog/no_schema_error'
require 'goog/no_spreadsheet_error'
require 'goog/sheet_record_collection'
require 'goog/no_session_error'
require 'goog/sheet_not_found_error'

class Goog::SheetRecord
  include Goog::ColumnToLetter
  attr :row_values
  attr :key_values
  attr :row_num
  attr :spreadsheet_id
  attr :sheet
  attr :changes
  attr :major_dimension

  def initialize(explicit_schema: nil, row_values: nil, row_num: nil, sheet: nil, spreadsheet_id: nil, major_dimension: nil)    
    @row_values = row_values
    @row_num = row_num
    @spreadsheet_id = spreadsheet_id
    @sheet = sheet
    @major_dimension = major_dimension
    @key_values = {}
    if explicit_schema
      @schema = explicit_schema
    end
    if @sheet.is_a?(String)
      @sheet = Goog::Services.sheets.get_sheet_by_name(@spreadsheet_id, @sheet)
    end
  end

  def schema
    if @schema.nil?
      raise Goog::NoSessionError if Goog::Services.session.nil?
      @schema = Goog::Services.session.get_schema(spreadsheet_id: @spreadsheet_id, sheet: @sheet)
      if @schema.nil?
        @schema = self.load_schema
      end
      raise Goog::NoSchemaError unless @schema
    end
    @schema
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

  def load_schema
    range = if @major_dimension == :rows
      '1:1'
    else
      'A:A'
    end
    values = Goog::Services.sheets.get_range(@spreadsheet_id, range, sheet: @sheet, value_render_option: :unformatted_value, major_dimension: @major_dimension)
    schema = self.class.create_schema(values[0])
    Goog::Services.session.set_schema(spreadsheet_id: @spreadsheet_id, sheet: @sheet, schema: schema)
    schema
  end

  def changes
    @changes ||= []
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
      result.push (index+1).to_s
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
    value = '' if value.nil?
    previous_value = self.get_row_value(key)
    return value if previous_value.blank? and value.blank?
    if previous_value != value
      self.changes.push [key, previous_value, value]
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
    self.changes.map { |key, previous_value, new_value|
      "#{key} #{previous_value} => #{new_value}"
    }.join(', ')
  end
  def changed?
    self.changes.size > 0
  end

  def row_values
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
    raise Goog::NoSpreadsheetError unless self.spreadsheet_id
    if self.row_num
      Goog::Services.sheets.write_changes(self.spreadsheet_id, self, major_dimension: self.major_dimension)
    else
      sheet_name = (@sheet.is_a?(String) ? @sheet : @sheet.properties.title)
      Goog::Services.sheets.append_range(self.spreadsheet_id, 
                                         "#{sheet_name}!A1:A1",
                                         [self.row_values])
    end
  end

  def self.find(spreadsheet_id, range: nil, sheet: nil, value_render_option: :unformatted_value, major_dimension: :rows)
    if sheet.is_a?(String)
      sheet = Goog::Services.sheets.get_sheet_by_name(spreadsheet_id, sheet) || (raise Goog::SheetNotFoundError)
    end
    sheet ||= Goog::Services.sheets.get_sheets(spreadsheet_id).first
    if values = Goog::Services.sheets.get_range(spreadsheet_id, range, sheet: sheet, major_dimension: major_dimension)
      self.from_range_values(values: values, spreadsheet_id: spreadsheet_id, sheet: sheet, major_dimension: major_dimension)
    else
      []
    end
  end

  def self.to_collection(values: nil, spreadsheet_id:, sheet:, multiple_sheet_values: nil)
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
    # if self.schema.nil?
    #   self.schema = self.create_schema(values[0])
    # end
    values[1..-1].map.with_index do |row, index|
      new(row_values: row, 
          row_num: index+2, 
          spreadsheet_id: spreadsheet_id, 
          sheet: sheet,
          major_dimension: major_dimension)
    end
  end

  def _rounded(attribute)
    (self.send(attribute) * 100).round / 100.0
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
    if @row_values
      if mid[-1] == '='
        mid = mid[0..-2]
      end
      self.schema.member?(mid.to_sym)
    else
      true
    end
  end

  def method_missing(mid, *args)
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

  GOOGLE_EPOCH=Date.new(1899,12,30)

  def self.google_integer_to_date(google_date)
    GOOGLE_EPOCH.advance(days: google_date)
  end

  def self.date_attribute(key)
    self.class_eval <<-RUBY
    def #{key}
      value = self.get_row_value(:#{key})
      return nil unless value.present?
      if self.looks_like_date?(value)
        begin
          value = Date.strptime(value, '%m/%d/%Y')
        rescue ArgumentError
        end
      elsif value.is_a?(Integer)
        value = self.class.google_integer_to_date(value)
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

  # This is separate to keep the special characters straight during class_eval.
  def parse_hyperlink_value(value)
    if value =~ /=HYPERLINK\("([^"]+)","([^"]+)"\)/
      value, link = $2, $1
    else
      value
    end
  end    

  def self.hyperlink_attribute(key)
    self.class_eval <<-RUBY
      def #{key}_parts
        value = self.get_row_value(:#{key})
        return nil unless value.present?
        self.parse_hyperlink_value(value)
      end
      def #{key}
        value, link = self.#{key}_parts
        value
      end
      def #{key}=(new_value)
        value, link = self.#{key}_parts
        self.set_#{key}(new_value, link)
      end
      def #{key}_hyperlink
        self.get_row_value(:#{key})
      end
      def #{key}_url
        value, link = self.#{key}_parts
        link
      end
      def #{key}_url=(new_url)
        value, link = self.#{key}_parts
        self.set_#{key}(value, new_url)
      end
      def #{key}_hyperlinked?
        value = self.get_row_value(:#{key})
        return false if value.nil?
        value.include?('HYPERLINK')
      end
      def set_#{key}(value, link=nil)
        if link
          link_value = '=HYPERLINK("' + link + '","' + value + '")'
          self.set_row_value(:#{key}, link_value)
        else
          self.set_row_value(:#{key}, value)
        end
      end
      def unlink_#{key}
        self.set_#{key}(self.#{key})
      end
    RUBY
  end

  protected

  def looks_like_date?(value)
    value.is_a?(String) && (value =~ /\d+\/\d+\/\d+/)
  end  
end