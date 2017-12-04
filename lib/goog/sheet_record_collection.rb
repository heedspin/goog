class Goog::SheetRecordCollection
  attr :records

  def initialize(records)
    @records = records
  end

  def first
    @records.first
  end

  def size
    @records.size
  end

  def sort(field, reverse: false)
    @records.sort_by! { |r| r.send(field) }
    @records.reverse! if reverse
  end

  def find(fields=nil, &block)
    @records.each do |t|
      if block_given?
        if yield(t)
          return t
        end
      elsif fields and fields.all? { |key, value| t.send(key) == value }
        return t
      end
    end
    nil
  end

  def find_all(fields = nil, &block)
    result = []
    @records.each do |t|
      if block_given?
        if yield(t)
          result.push t
        end
      elsif fields.nil? or fields.all? { |key, value| t.send(key) == value }
        result.push t
      end
    end
    result
  end
end