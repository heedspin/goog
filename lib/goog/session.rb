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
end