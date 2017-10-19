class Goog::SheetRecord
  attr :schema
  attr :row_values
  attr :row_num
  attr :sheet
  include Goog::SpreadsheetUtils

  def initialize(schema:, row_values:, row_num:, sheet:)
    @schema = schema
    @row_values = row_values
    @row_num = row_num
    @sheet = sheet
  end

  def get_row_value(key)
    if index = @schema[key]
      @row_values[index]
    else
      nil
    end
  end

  def set_row_value(key, value)
    if index = @schema[key]
      @row_values[index] = value
    else
      raise ArgumentError.new("#{key} is not in schema")
    end
  end

  def method_missing(mid, *args)
    if mid[-1] == '='
      mid = mid[0..-2].to_sym
      raise NoMethodError.new("Unknown method #{mid}=") unless @schema.member?(mid)
      self.set_row_value(mid, args[0])
    else
      mid = mid.to_sym
      raise NoMethodError.new("Unknown method #{mid}") unless @schema.member?(mid)
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

  def self.from_range_values(values:, sheet:)
    schema = self.create_schema(values[0])
    values[1..-1].map.with_index do |row, index|
      new(schema: schema, row_values: row, row_num: index+2, sheet: sheet)
    end
  end


end