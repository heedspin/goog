class Goog::SheetRecord
  attr :schema
  attr :row_values
  include Goog::SpreadsheetUtils

  def initialize(schema:, row_values:)
    @schema = schema
    @row_values = row_values
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
      self.set_row_value(mid[0..-2].to_sym, args[0])
    else
      self.get_row_value(mid.to_sym)
    end
  end

  def self.create_schema(header_row)
    result = {}
    header_row.each_with_index do |value, index|
      result[value.parameterize.underscore.to_sym] = index
    end
    result
  end
end