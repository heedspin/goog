# https://developers.google.com/drive/v3/web/about-sdk
require 'plutolib/logger_utils'
require 'google/apis/drive_v3'

module Goog::DriveUtils
  def current_drive
    if @current_drive.nil?
      raise "No authorizer established" unless self.current_authorizer.present?
      @current_drive = Google::Apis::DriveV3::DriveService.new
      @current_drive.authorization = self.current_authorizer
    end
    @current_drive
  end

  # Returns nil on failure.  Permission id on success.
  def add_writer_permission(file_id:, email_address:)
    # https://developers.google.com/drive/v3/web/manage-sharing
    permission = { type: 'user', role: 'writer', email_address: email_address }
    goog_retries do
      result = self.current_drive.create_permission(file_id,
                                                    permission, 
                                                    fields: 'id')
      return result.id
    end
  end

  def get_files_in_folder(folder_id)
    goog_retries do
      result = self.current_drive.list_files(corpora: 'user', q: ["\"#{folder_id}\" in parents"])
      return result.files
    end
  end

  def delete_files_containing(containing, drive: nil)
    drive ||= self.current_drive
    # Delete the mess we created.
    drive.list_files(corpora: 'user', q: ["name contains \"#{containing}\""]).files.each do |file|
      begin
        drive.delete_file(file.id)
        log "Deleted #{file.name}"
      rescue Google::Apis::ClientError => e
        log_error "Failed to delete #{file.name}"
      end
    end
  end

  def self.included base
    base.class_eval do
      include Plutolib::LoggerUtils
      include Goog::AuthorizerUtils
    end
  end
end