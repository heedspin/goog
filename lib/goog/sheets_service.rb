# https://developers.google.com/sheets/api/
require 'google/apis/sheets_v4'
require 'plutolib/logger_utils'

class Goog::SheetsService
  include Plutolib::LoggerUtils
  include Goog::Retry
  include Goog::ColumnToLetter
  attr_accessor :sheets
  def initialize(authorization)
    @sheets = Google::Apis::SheetsV4::SheetsService.new
    @sheets.authorization = authorization
  end
  def create_spreadsheet(title:, writer_emails: nil)
    spreadsheet = nil
    goog_retries do
      spreadsheet = @sheets.create_spreadsheet
    end
    if self.rename_spreadsheet(spreadsheet_id: spreadsheet.spreadsheet_id, title: title)
      if writer_emails
        if writer_emails.is_a?(String)
          writer_emails = [writer_emails]
        end
        writer_emails.each do |email_address|
          Goog::Services.drive.add_writer_permission(file_id: spreadsheet.spreadsheet_id, email_address: email_address)
        end
      end
    end
    spreadsheet
  end

  # Returns true on success.
  def rename_spreadsheet(spreadsheet_id:, title:)
    requests = []
    requests.push({
      update_spreadsheet_properties: {
        properties: {title: title},
        fields: 'title'
      }
    })
    goog_retries do
      response = @sheets.batch_update_spreadsheet(spreadsheet_id, 
                                                                      {requests: requests}, 
                                                                      {})
      response.spreadsheet_id.present?
    end
  end

  # Returns true on success.
  def turn_off_filters(spreadsheet_id:, sheet:)
    requests = []
    requests.push({
      clear_basic_filter: {
        sheet_id: sheet.properties.sheet_id
      }
    })
    goog_retries do
      response = @sheets.batch_update_spreadsheet(spreadsheet_id, 
                                                                      {requests: requests}, 
                                                                      {})
      response.spreadsheet_id.present?
    end
  end

  def get_spreadsheet(spreadsheet_id)
    goog_retries do
      @sheets.get_spreadsheet(spreadsheet_id, fields: 'sheets.properties')
    end
  end

  def get_sheets(spreadsheet_id)
    goog_retries do
      result = @sheets.get_spreadsheet(spreadsheet_id, fields: 'sheets.properties')
      result.sheets
    end
  end

  def get_sheet_by_name(spreadsheet_id, name)
    self.get_sheets(spreadsheet_id).each do |sprop|
      if sprop.properties.title.downcase == name.downcase
        return sprop
      end
    end
    nil
  end

  def get_range(spreadsheet_id, range=nil, sheet: nil, value_render_option: :unformatted_value, major_dimension: :rows)
    spreadsheet_id = spreadsheet_id.id if spreadsheet_id.is_a?(Google::Apis::DriveV3::File)
    if range.nil?
      if sheet
        sheet = self.get_sheet_by_name(spreadsheet_id, sheet) if sheet.is_a?(String)
      else
        sheet = self.get_sheets(spreadsheet_id).first
      end
      range = self.get_sheet_range(sheet)
    end
    goog_retries do
      result = @sheets.get_spreadsheet_values(spreadsheet_id, 
                                              range, 
                                              value_render_option: value_render_option,
                                              major_dimension: major_dimension)
      result.values
    end
  end

  def get_multiple_ranges(spreadsheet_id, ranges: nil, sheets: nil, value_render_option: :unformatted_value)
    sheets ||= self.get_sheets(spreadsheet_id)
    ranges ||= sheets.map { |s| self.get_sheet_range(s) }.compact
    raise "No ranges specified" unless ranges
    goog_retries do
      result = @sheets.batch_get_spreadsheet_values(spreadsheet_id, 
                                                    ranges: ranges, 
                                                    value_render_option: value_render_option)
      result.value_ranges
    end
  end

  def clear_range(spreadsheet_id, range)
    goog_retries do
      request_body = Google::Apis::SheetsV4::ClearValuesRequest.new
      result = @sheets.clear_values(spreadsheet_id, range, request_body)
    end
  end

  # Returns hash of tab names to value arrays.
  def get_multiple_sheet_values(spreadsheet_id, sheets: nil, value_render_option: :unformatted_value)
    sheets ||= self.get_sheets(spreadsheet_id)
    value_ranges = self.get_multiple_ranges(spreadsheet_id, 
                                            sheets: sheets, 
                                            value_render_option: value_render_option)
    result = {}
    value_ranges.each do |value_range|
      tab_name = self.get_range_parts(value_range.range)[0]
      result[tab_name] = value_range.values
    end
    result
  end

  def write_rows(sheet:, spreadsheet_id:, start_row:, values:)
    range = "#{sheet.properties.title}!A#{start_row}"
    goog_retries do
      @sheets.update_spreadsheet_value(spreadsheet_id,
                                                           range,
                                                           { values: values },
                                                           value_input_option: 'USER_ENTERED')
    end
    true
  end

  def write_range(spreadsheet_id, range, values)
    goog_retries do
      @sheets.update_spreadsheet_value(spreadsheet_id,
                                                           range,
                                                           { values: values },
                                                           value_input_option: 'USER_ENTERED')
    end
  end

  def append_range(spreadsheet_id, range, values, major_dimension: :rows)
    goog_retries do
      @sheets.append_spreadsheet_value(spreadsheet_id,
                                                           range,
                                                           { values: values, major_dimension: major_dimension },
                                                           value_input_option: 'USER_ENTERED')
    end
  end

  def batch_write_ranges(spreadsheet_id, data, major_dimension: :rows)    
    goog_retries do
      @sheets.batch_update_values(spreadsheet_id, 
                                                      { value_input_option: 'USER_ENTERED', data: data, major_dimension: major_dimension },
                                                      { })
    end
    true
  end

  def write_changes(spreadsheet_id, record, major_dimension: :rows)
    data = []
    record.changes.each do |key, previous_value, value|
      data.push({ range: record.a1(key), values: [[value]] })
    end
    self.batch_write_ranges(spreadsheet_id, data, major_dimension: major_dimension)
  end

  def insert_empty_rows(sheet:, spreadsheet_id:, start_row:, num_rows: 1)
    requests = []
    requests.push({
      insert_dimension: {
        range: {
          sheet_id: sheet.properties.sheet_id,
          dimension: 'ROWS',
          start_index: start_row - 1,
          end_index: start_row + num_rows - 1
        },
        inherit_before: false
      }
    })
    goog_retries do
      @sheets.batch_update_spreadsheet(spreadsheet_id, 
                                                           {requests: requests}, 
                                                           {})
    end
    true
  end

  def move_row_between_sheets(spreadsheet_id:, source_sheet:, source_row_num:, destination_sheet:, destination_row_num:)
    self.insert_empty_rows(spreadsheet_id: spreadsheet_id, 
                           sheet: destination_sheet, 
                           start_row: destination_row_num,
                           num_rows: 1)
    self.cut_paste_rows(spreadsheet_id: spreadsheet_id,
                        source_sheet: source_sheet,
                        source_row_num: source_row_num,
                        destination_sheet: destination_sheet,
                        destination_row_num: destination_row_num)
    self.delete_rows(spreadsheet_id: spreadsheet_id,
                     sheet: source_sheet,
                     row_num: source_row_num,
                     num_rows: 1)
  end

  def cut_paste_rows(spreadsheet_id:, source_sheet:, source_row_num:, num_rows:1, destination_sheet:,destination_row_num:)
    requests = [ {
      cut_paste: {
        source: {
          sheet_id: source_sheet.properties.sheet_id,
          start_row_index: source_row_num-1,
          end_row_index: source_row_num-1+num_rows,
          start_column_index: 0,
          end_column_index: destination_sheet.properties.grid_properties.column_count
        },
        destination: {
          sheet_id: destination_sheet.properties.sheet_id,
          row_index: destination_row_num-1,
          column_index: 0
        },
        paste_type: 'PASTE_NORMAL'
      }
    } ]
    goog_retries do
      @sheets.batch_update_spreadsheet(spreadsheet_id, {requests: requests}, {})
    end
    true
  end

  # Thanks: https://stackoverflow.com/questions/12913874/is-it-possible-to-use-the-google-spreadsheet-api-to-add-a-comment-in-a-cell
  def add_note(note:, spreadsheet_id:, sheet:, row_num:, num_rows:1, column_num:, num_cols: 1)
    requests = [ {
      repeat_cell: {
        range: {
          sheet_id: sheet.properties.sheet_id,
          start_row_index: row_num-1,
          end_row_index: row_num-1+num_rows,
          start_column_index: column_num-1,
          end_column_index: column_num-1+num_cols
        },
        cell: { note: note },
        fields: 'note'
      }
    } ]
    goog_retries do
      @sheets.batch_update_spreadsheet(spreadsheet_id, {requests: requests}, {})
    end
    true
  end

  # Mind that row_num is NOT an index.  Starts counting with 1!
  def delete_rows(spreadsheet_id:, sheet:, row_num:, num_rows: 1)
    requests = [ {
      delete_dimension: {
        range: {
          sheet_id: sheet.properties.sheet_id,
          dimension: 'ROWS',
          start_index: row_num-1,
          end_index: row_num-1+num_rows
        }
      }
    } ]
    goog_retries do
      @sheets.batch_update_spreadsheet(spreadsheet_id, {requests: requests}, {})
    end
    true
  end

  def copy_spreadsheet(spreadsheet_id, new_name: nil, writer_emails: nil, destination_folder_id: nil)
    new_file = nil
    goog_retries do
      new_file = Goog::Services.drive.drive.copy_file(spreadsheet_id)
    end
    if new_name
      # Renames the spreadsheet
      requests = []
      requests.push({
        update_spreadsheet_properties: {
          properties: {title: new_name},
          fields: 'title'
        }
      })
      goog_retries do
        @sheets.batch_update_spreadsheet(new_file.id, 
                                                             {requests: requests},
                                                             {})
      end
    end
    if writer_emails
      Goog::Services.drive.add_writer_permission(file_id: new_file.id, email_address: writer_emails) || (return false)
    end
    if destination_folder_id
      Goog::Services.drive.add_file_to_folder(file_id: new_file.id, folder_id: destination_folder_id) || (return false)
    end
    new_file
  end

  # Returns true if it moved.  false if not moved.  exception on error.
  def move_sheet(sheet_id, destination_folder_id)
    previous_parents = self.current_drive.get_file(sheet_id, fields: 'parents').parents
    if previous_parents.include?(destination_folder_id)
      false
    else
      goog_retries do
        Goog::Services.drive.update_file(sheet_id,
                                         add_parents: destination_folder_id,
                                         remove_parents: previous_parents.join(','),
                                         fields: 'id, parents')
      end
      true
    end
  end

  def get_sheet_range(sheet_properties, columns: nil)
    return nil unless sheet_properties
    column_count = columns ? columns : sheet_properties.properties.grid_properties.column_count
    "#{sheet_properties.properties.title}!A1:#{self.column_to_letter(column_count)}#{sheet_properties.properties.grid_properties.row_count}"
  end

  # Returns [Sheet, Range1, Range2]
  def get_range_parts(range)
    if range =~ /([' a-zA-Z]+!)?([A-Z]+\d+):([A-Z]+\d+)/
      sheet = $1
      range1 = $2
      range2 = $3
      if sheet
        sheet = sheet[0..sheet.size-2]
        if sheet =~ /'(.+)'/
          sheet = $1
        end
      end
      return [sheet, range1, range2]
    end
    nil
  end

  def self.included base
    base.class_eval do
      include Goog::DriveUtils
      include Goog::DateUtils
    end
  end
end