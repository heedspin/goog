require 'goog/session_not_open'

class Goog::Session
  def self.start(&block)
    session = new
    begin
      session.open
      Goog::Services.session = session
      yield if block_given?
    ensure
      session.close if block_given?
    end
  end
  def self.end
    Goog::Services.session.try(:close)
    Goog::Services.session = nil
  end

  def open
    unless @opened
      @schemas = {}
      @opened = true
    end
    @opened
  end

  def close
    if @opened
      @schemas = nil
      @opened = false
    end
    true
  end

  def get_schema(spreadsheet_id:, sheet:)
    raise Goog::SessionNotOpen unless @opened
    @schemas[[spreadsheet_id, sheet.properties.sheet_id]]
  end
  def set_schema(spreadsheet_id:, sheet:, schema:)
    raise Goog::SessionNotOpen unless @opened
    @schemas[[spreadsheet_id, sheet.properties.sheet_id]] = schema
  end

  def enable_profiling
    @profiling = true
    @profiling_context = []
    @profiling_history = []
    @profiling_event_type_counts = Hash.new(0)
  end

  def profiling_enabled?
    @profiling
  end

  def profile_context_push(name)
    return unless @profiling
    @profiling_history.push [@profiling_context.size, 'Begin Context: ' + @profiling_context.join(' / ')]
    @profiling_context.push(name)
  end
  def profile_context_pop
    return unless @profiling
    @profiling_context.pop
    @profiling_history.push [@profiling_context.size, 'End Context: ' + @profiling_context.join(' / ')]
  end
  def profile_event(type, name)
    return unless @profiling
    @profiling_history.push [@profiling_context.size, "#{type}: #{name}"]
    @profiling_event_type_counts[type] += 1
  end
  def profile_dump(path)
    return unless @profiling
    File.open(path, 'w+') do |output|
      output.puts "Event Counts:"
      @profiling_event_type_counts.to_a.sort_by(&:last).each do |event_type, count|
        output.puts "  #{event_type}: #{count}"
      end
      @profiling_history.each do |depth, description|
        output.puts ('  ' * depth) + description
      end
    end
  end
end