# https://developers.google.com/drive/v3/web/about-sdk
require 'plutolib/logger_utils'
require 'google/apis/drive_v3'

class Goog::DriveService
  include Plutolib::LoggerUtils
  include Goog::Retry
  
  attr_accessor :drive
  def initialize(authorization)
    @drive = Google::Apis::DriveV3::DriveService.new
    @drive.authorization = authorization
  end

  # Returns nil on failure.  Permission id on success.
  def add_permissions(file_id:, writer_emails: nil, owner_emails: nil)
    result = []
    if writer_emails
      result.append self.add_writer_permission(file_id: file_id, email_address: writer_emails) 
    end
    if owner_emails
      result.append self.add_owner_permission(file_id: file_id, email_address: owner_emails) 
    end
    result
  end

  # Returns nil on failure.  Permission id on success.
  def add_writer_permission(file_id:, email_address:)
    # https://developers.google.com/drive/v3/web/manage-sharing
    writer_emails = [email_address] if email_address.is_a?(String)
    result = []
    writer_emails.each do |email|
      permission = { type: 'user', role: 'writer', email_address: email }
      goog_retries do
        result.push @drive.create_permission(file_id, permission, fields: 'id')
      end
    end    
    return result.size == 1 ? result.first : result
  end

  # Returns nil on failure.  Permission id on success.
  def add_owner_permission(file_id:, email_address:)
    # https://developers.google.com/drive/v3/web/manage-sharing
    owner_emails = [email_address] if email_address.is_a?(String)
    result = []
    owner_emails.each do |email|
      permission = { type: 'user', role: 'owner', email_address: email }
      goog_retries do
        result.push @drive.create_permission(file_id, permission, fields: 'id', transfer_ownership: true)
      end
    end    
    return result.size == 1 ? result.first : result
  end
  def get_permission(file, permission, fields: 'display_name, email_address')
    goog_retries do
      @drive.get_permission(file.id, permission.id, fields: fields)
    end
  end

  def list_permissions(file)
    goog_retries do
      @drive.list_permissions(file.id).permissions
    end
  end

  def build_drive_utils_query(query, parent_folder_id: nil, file_type: nil)
    if parent_folder_id
      query.push "parents in '#{parent_folder_id}'"
    end
    case file_type
    when :folder
      query.push "mimeType = 'application/vnd.google-apps.folder'"
    when :file
      query.push "mimeType != 'application/vnd.google-apps.folder'"
    end
  end

  def fetch_all(&block)
    @drive.fetch_all(items: :files) do |page_token|
      goog_retries do 
        @drive.list_files(corpora: 'user', include_team_drive_items: false)
      end
    end
  end

  def get_files_containing(containing, parent_folder_id: nil, file_type: :file)
    query = [] 
    query.push("name contains '#{containing}'") if containing.present?
    self.build_drive_utils_query(query, parent_folder_id: parent_folder_id, file_type: file_type)
    goog_retries do
      result = @drive.list_files(corpora: 'user', include_team_drive_items: false, q: query.join(' and '))
      return result.files
    end
  end

  def get_folders_by_name(name, parent_folder_id: nil)
    self.get_files_by_name(name, parent_folder_id: parent_folder_id, file_type: :folder)
  end

  def get_folders_containing(containing, parent_folder_id: nil)
    self.get_files_containing(containing, parent_folder_id: parent_folder_id, file_type: :folder)
  end

  # https://developers.google.com/drive/v3/web/search-parameters
  def get_files_by_name(name, parent_folder_id: nil, file_type: :file)
    query = ["name = '#{name}'"]
    self.build_drive_utils_query(query, parent_folder_id: parent_folder_id, file_type: file_type)
    goog_retries do
      result = @drive.list_files(corpora: 'user', include_team_drive_items: false, q: query.join(' and '))
      return result.files
    end
  end

  def create_folder(name, parent_folder_id: nil, writer_emails: nil, owner_emails: nil)
    file_metadata = {
      name: name,
      mime_type: 'application/vnd.google-apps.folder'
    }
    if parent_folder_id
      file_metadata[:parents] = [parent_folder_id]
    end
    file_id = nil
    goog_retries do
      file_id = @drive.create_file(file_metadata, fields: 'id').try(:id)
    end
    self.add_permissions(file_id: file_id, writer_emails: writer_emails, owner_emails: owner_emails)
    file_id
  end

  def folder_exists?(folder_id)
    begin
      !@drive.get_file(folder_id).nil?
    rescue Google::Apis::ClientError
      return false
    end
  end

  def get_files_in_folder(parent_folder_id)
    goog_retries do
      result = @drive.list_files(corpora: 'user', q: ["\"#{parent_folder_id}\" in parents"])
      return result.files
    end
  end

  def add_file_to_folder(file_id:, folder_id:)
    previous_parents = @drive.get_file(file_id, fields: 'parents').parents.join(',')
    goog_retries do
      @drive.update_file(file_id,
                         add_parents: folder_id,
                         remove_parents: previous_parents,
                         fields: 'id, parents')
    end
    true
  end

  def move_file(file_id: , parent_folder_id:)
    file_id = file_id.id if file_id.is_a?(Google::Apis::DriveV3::File)
    previous_parents = @drive.get_file(file_id, fields: 'parents').parents
    if previous_parents.include?(parent_folder_id)
      false
    else
    goog_retries(profile_type: 'Drive#move_file') do
        @drive.update_file(file_id,
                           add_parents: parent_folder_id,
                           remove_parents: previous_parents.join(','),
                           fields: 'id, parents')
      end
      true
    end
  end

  def delete_files_containing(containing, parent_folder_id: nil, file_type: :file)
    self.get_files_containing(containing, parent_folder_id: parent_folder_id, file_type: file_type).each do |file|
      begin
        goog_retries do
          @drive.delete_file(file.id)
        end
        log "Deleted #{file.name}"
      rescue Google::Apis::ClientError => e
        log_error "Failed to delete #{file.name}"
      end
    end
  end

  def rename_file(file_id, new_name)
    file_id = file_id.id if file_id.is_a?(Google::Apis::DriveV3::File)
    file_metadata = {
      name: new_name,
    }
    goog_retries do
      @drive.update_file(file_id, file_metadata, fields: 'id').try(:id)
    end
  end

  # Google Drive MIME Types: https://developers.google.com/drive/v3/web/mime-types
  def upload_odt_to_doc(existing_file_id: nil, file_name: nil, folder_id: nil, path_to_odt:, writer_emails: nil, owner_emails: nil)
    file_id = nil
    file_metadata = {
      name: file_name,
      mime_type: 'application/vnd.google-apps.document'
    }
    if existing_file_id
      goog_retries do
        file_id = @drive.update_file(existing_file_id,
                                     file_metadata,
                                     fields: 'id',
                                     upload_source: path_to_odt,
                                     content_type: 'application/vnd.oasis.opendocument.text').try(:id)
      end
    else
      goog_retries do
        file_id = @drive.create_file(file_metadata,
                                     fields: 'id',
                                     upload_source: path_to_odt,
                                     content_type: 'application/vnd.oasis.opendocument.text').try(:id)
      end
      if folder_id
        self.add_file_to_folder(file_id: file_id, folder_id: folder_id) || (return false)
      end
    end
    if file_id      
      self.add_permissions(file_id: file_id, writer_emails: writer_emails, owner_emails: owner_emails)
    end
    file_id
  end

  # https://developers.google.com/drive/v3/web/manage-downloads
  def download_document_to_odt(file_id, download_path)
    @drive.export_file(file_id,
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

  def quote_sed_string(str)
    str.to_s.gsub('[','\[').gsub(']','\]').gsub("'", '\\\'')
  end

  # Returns new file id.
  # You do not need to specify shared folder and writer_emails.  If folder is shared, permissions are inherited.
  def create_doc_from_template(source_file_id:, destination_folder_id: nil, destination_file_name:, destination_file_id: nil, query_replace_map:, writer_emails: nil, owner_emails: nil)
    file_id = nil
    Dir.mktmpdir do |dir|
      odt_file = File.join(dir, 'doc.odt')
      content_file = File.join(dir, 'content.xml')
      self.download_document_to_odt(source_file_id, odt_file) || (return false)
      self.drive_shell_command("unzip -d #{dir} #{odt_file} content.xml") || (return false)
      # Do query replace
      query_replace_map.each do |from, to|
        self.drive_shell_command("sed -i 's/#{quote_sed_string(from)}/#{quote_sed_string(to)}/g' #{content_file}") || (return false)
      end
      self.drive_shell_command("cd #{dir} ; zip #{odt_file} content.xml") || (return false)
      file_id = self.upload_odt_to_doc(existing_file_id: destination_file_id, file_name: destination_file_name, path_to_odt: odt_file, writer_emails: writer_emails, owner_emails: owner_emails)
    end
    if destination_folder_id
      self.add_file_to_folder(file_id: file_id, folder_id: destination_folder_id) || (return false)
    end
    return file_id
  end

  def self.included base
    base.class_eval do
      include Plutolib::LoggerUtils
      include Goog::Retry
    end
  end
end