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
    writer_emails = [email_address] if email_address.is_a?(String)
    result = []
    writer_emails.each do |email|
      permission = { type: 'user', role: 'writer', email_address: email }
      goog_retries do
        result.push self.current_drive.create_permission(file_id, permission, fields: 'id')
      end
    end    
    return result.size == 1 ? result.first : result
  end

  def get_files_containing(containing, parent_folder_id: nil)
    query = ["name contains \"#{containing}\""]
    if parent_folder_id
      query.push "\"#{parent_folder_id}\" in parents"
    end
    goog_retries do
      result = self.current_drive.list_files(corpora: 'user', q: query)
      return result.files
    end
  end

  # https://developers.google.com/drive/v3/web/search-parameters
  def get_files_by_name(name, parent_folder_id: nil)
    query = ["name = \"#{name}\""]
    if parent_folder_id
      query.push "\"#{parent_folder_id}\" in parents"
    end
    goog_retries do
      result = self.current_drive.list_files(corpora: 'user', q: query)
      return result.files
    end
  end

  def create_folder(name, parent_folder_id: nil, writer_emails: nil)
    file_metadata = {
      name: name,
      mime_type: 'application/vnd.google-apps.folder'
    }
    if parent_folder_id
      file_metadata[:parents] = [parent_folder_id]
    end
    file_id = nil
    goog_retries do
      file_id = self.current_drive.create_file(file_metadata, fields: 'id').try(:id)
    end
    if writer_emails
      self.add_writer_permission(file_id: file_id, email_address: writer_emails)
    end
    file_id
  end

  def get_files_in_folder(parent_folder_id)
    goog_retries do
      result = self.current_drive.list_files(corpora: 'user', q: ["\"#{parent_folder_id}\" in parents"])
      return result.files
    end
  end

  def add_file_to_folder(file_id:, destination_folder_id:)
    previous_parents = self.current_drive.get_file(file_id, fields: 'parents').parents.join(',')
    goog_retries do
      self.current_drive.update_file(file_id,
                                     add_parents: destination_folder_id,
                                     remove_parents: previous_parents,
                                     fields: 'id, parents')
    end
    true
  end    

  def delete_files_containing(containing)
    self.get_files_containing(containing).each do |file|
      begin
        self.current_drive.delete_file(file.id)
        log "Deleted #{file.name}"
      rescue Google::Apis::ClientError => e
        log_error "Failed to delete #{file.name}"
      end
    end
  end

  # Google Drive MIME Types: https://developers.google.com/drive/v3/web/mime-types
  def upload_odt_to_doc(existing_file_id: nil, file_name: nil, path_to_odt: , writer_emails: nil)
    file_id = nil
    file_metadata = {
      name: file_name,
      mime_type: 'application/vnd.google-apps.document'
    }
    if existing_file_id
      goog_retries do
        file_id = self.current_drive.update_file(existing_file_id,
                                                 file_metadata,
                                                 fields: 'id',
                                                 upload_source: path_to_odt,
                                                 content_type: 'application/vnd.oasis.opendocument.text').try(:id)
      end
    else
      goog_retries do
        file_id = self.current_drive.create_file(file_metadata,
                                                 fields: 'id',
                                                 upload_source: path_to_odt,
                                                 content_type: 'application/vnd.oasis.opendocument.text').try(:id)
      end
    end
    if writer_emails and file_id
      self.add_writer_permission(file_id: file_id, email_address: writer_emails)
    end
    file_id
  end

  # https://developers.google.com/drive/v3/web/manage-downloads
  def download_document_to_odt(file_id, download_path)
    self.current_drive.export_file(file_id,
                                   'application/vnd.oasis.opendocument.text',
                                    download_dest: download_path)
    true
  end

  def drive_shell_command(command)
    output = `#{command}`
    status = $?
    log command + "\n" + output
    return status.exitstatus == 0
  end

  # Returns new file id.
  # You do not need to specify shared folder and writer_emails.  If folder is shared, permissions are inherited.
  def create_doc_from_template(source_file_id:, destination_folder_id: nil, destination_file_name:, destination_file_id: nil, query_replace_map:, writer_emails: nil)
    file_id = nil
    Dir.mktmpdir do |dir|
      odt_file = File.join(dir, 'doc.odt')
      content_file = File.join(dir, 'content.xml')
      self.download_document_to_odt(source_file_id, odt_file) || (return false)
      self.drive_shell_command("unzip -d #{dir} #{odt_file} content.xml") || (return false)
      # Do query replace
      query_replace_map.each do |from, to|
        from = from.gsub('[','\[').gsub(']','\]')
        self.drive_shell_command("sed -i 's/#{from}/#{to}/g' #{content_file}") || (return false)
      end
      self.drive_shell_command("cd #{dir} ; zip #{odt_file} content.xml") || (return false)
      file_id = self.upload_odt_to_doc(existing_file_id: destination_file_id, file_name: destination_file_name, path_to_odt: odt_file, writer_emails: writer_emails)
    end
    if destination_folder_id
      self.add_file_to_folder(file_id: file_id, destination_folder_id: destination_folder_id) || (return false)
    end
    if writer_emails
      self.add_writer_permission(file_id: file_id, email_address: writer_emails)
    end
    return file_id
  end

  def self.included base
    base.class_eval do
      include Plutolib::LoggerUtils
      include Goog::Retry
      include Goog::AuthorizerUtils
    end
  end
end