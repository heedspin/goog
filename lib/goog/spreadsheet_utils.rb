# https://developers.google.com/sheets/api/
require 'google/apis/sheets_v4'

module Goog::SpreadsheetUtils
  def create_spreadsheet(title:, writer_emails: nil)
    spreadsheet = nil
    goog_retries do
      spreadsheet = self.current_sheets_service.create_spreadsheet
    end
    if self.rename_spreadsheet(spreadsheet_id: spreadsheet.spreadsheet_id, title: title)
      if writer_emails
        if writer_emails.is_a?(String)
          writer_emails = [writer_emails]
        end
        writer_emails.each do |email_address|
          self.add_writer_permission(file_id: spreadsheet.spreadsheet_id, email_address: email_address)
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
      response = self.current_sheets_service.batch_update_spreadsheet(spreadsheet_id, 
                                                                      {requests: requests}, 
                                                                      {})
      response.spreadsheet_id.present?
    end
  end

  def current_sheets_service
    if @current_sheets_service.nil?
      raise "No authorizer established" unless self.current_authorizer.present?
      @current_sheets_service = Google::Apis::SheetsV4::SheetsService.new
      @current_sheets_service.authorization = self.current_authorizer
    end
    @current_sheets_service
  end

  def get_spreadsheet(spreadsheet_id)
    goog_retries do
      self.current_sheets_service.get_spreadsheet(spreadsheet_id,
                                                  fields: 'sheets.properties')
    end
  end

  def get_sheets(spreadsheet_id)
    goog_retries do
      result = self.current_sheets_service.get_spreadsheet(spreadsheet_id,
                                                           fields: 'sheets.properties')
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

  def get_range(spreadsheet_id, range)
    goog_retries do
      result = self.current_sheets_service.get_spreadsheet_values(spreadsheet_id, range)
      result.values
    end
  end

  def get_multiple_ranges(spreadsheet_id, sheets:)
    ranges = sheets.map { |s| self.get_sheet_range(s) }
    goog_retries do
      result = self.current_sheets_service.batch_get_spreadsheet_values(spreadsheet_id, ranges: ranges)
      result.value_ranges
    end
  end

  # Returns hash of tab names to value arrays.
  def get_multiple_sheet_values(spreadsheet_id, sheets:)
    value_ranges = self.get_multiple_ranges(spreadsheet_id, sheets: sheets)
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
      self.current_sheets_service.update_spreadsheet_value(spreadsheet_id,
                                                           range,
                                                           { values: values },
                                                           value_input_option: 'USER_ENTERED')
    end
    true
  end

  def write_range(spreadsheet_id, range, values)
    goog_retries do
      self.current_sheets_service.update_spreadsheet_value(spreadsheet_id,
                                                           range,
                                                           { values: values },
                                                           value_input_option: 'USER_ENTERED')
    end
  end

  def append_range(spreadsheet_id, range, values)
    goog_retries do
      self.current_sheets_service.append_spreadsheet_value(spreadsheet_id,
                                                           range,
                                                           { values: values },
                                                           value_input_option: 'USER_ENTERED')
    end
  end

  def batch_write_ranges(spreadsheet_id, data)    
    goog_retries do
      self.current_sheets_service.batch_update_values(spreadsheet_id, 
                                                      { value_input_option: 'USER_ENTERED', data: data },
                                                      {})
    end
    true
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
      self.current_sheets_service.batch_update_spreadsheet(spreadsheet_id, 
                                                           {requests: requests}, 
                                                           {})
    end
    true
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
      self.current_sheets_service.batch_update_spreadsheet(spreadsheet_id, {requests: requests}, {})
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
      self.current_sheets_service.batch_update_spreadsheet(spreadsheet_id, {requests: requests}, {})
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
      self.current_sheets_service.batch_update_spreadsheet(spreadsheet_id, {requests: requests}, {})
    end
    true
  end

  def copy_spreadsheet(spreadsheet_id, new_name: nil, writer_emails: nil, destination_folder_id: nil)
    new_file = self.current_drive.copy_file(spreadsheet_id)
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
        self.current_sheets_service.batch_update_spreadsheet(new_file.id, 
                                                             {requests: requests},
                                                             {})
      end
    end
    if writer_emails
      writer_emails = [writer_emails] if writer_emails.is_a?(String)
      writer_emails.each do |email_address|
        permission = { type: 'user', role: 'writer', email_address: email_address }
        goog_retries do
          self.current_drive.create_permission(new_file.id,
                                               permission,
                                               fields: 'id')
        end
      end
    end
    if destination_folder_id
      previous_parents = self.current_drive.get_file(new_file.id, fields: 'parents').parents.join(',')
      goog_retries do
        self.current_drive.update_file(new_file.id,
                                       add_parents: destination_folder_id,
                                       remove_parents: previous_parents,
                                       fields: 'id, parents')
      end
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
        self.current_drive.update_file(sheet_id,
                                       add_parents: destination_folder_id,
                                       remove_parents: previous_parents.join(','),
                                       fields: 'id, parents')
      end
      true
    end
  end

  # https://stackoverflow.com/questions/21229180/convert-column-index-into-corresponding-column-letter
  def column_to_letter(column)
    letter = ''
    while (column > 0)
      temp = (column - 1) % 26
      letter = (temp + 65).chr + letter
      column = (column - temp - 1) / 26
    end
    letter
  end
  def letter_to_column(letter)
    column = 0
    length = letter.length
    for i in 0..(length-1)
      column += (letter[i].ord - 64) * (26 ** (length - i - 1))
    end
    column
  end

  def get_sheet_range(sheet_properties, columns: nil)
    column_count = columns ? columns : sheet_properties.properties.grid_properties.column_count
    "#{sheet_properties.properties.title}!A1:#{self.column_to_letter(column_count)}#{sheet_properties.properties.grid_properties.row_count}"
  end

  # Returns [Sheet, Range1, Range2]
  def get_range_parts(range)
    if range =~ /([a-zA-Z]+!)?([A-Z]+\d+):([A-Z]+\d+)/
      if sheet = $1
        sheet = sheet[0..sheet.size-2]
      end
      return [sheet, $2, $3]
    end
    nil
  end

  def result_set_to_records(record_class, range_values)
    schema = record_class.create_schema(range_values[0])
    range_values[1..-1].map do |row|
      record_class.new(schema: schema, row_values: row)
    end
  end

  def self.included base
    base.class_eval do
      include Goog::DriveUtils
    end
  end
end