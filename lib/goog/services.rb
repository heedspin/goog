require 'plutolib/logger_utils'
# https://github.com/google/google-auth-library-ruby
module Goog::Services
  include Goog::Retry
  include Plutolib::LoggerUtils
  class << self
    attr_accessor :authorization
    attr_accessor :drive
    attr_accessor :sheets
    attr_accessor :session
  end

  # On behalf of domain user:
  # https://developers.google.com/identity/protocols/OAuth2ServiceAccount

  # Returns current authorizer
  def self.authorize_service_account(auth_file:, scope: Google::Apis::DriveV3::AUTH_DRIVE)
    self.authorization = Google::Auth::ServiceAccountCredentials.make_creds(json_key_io: File.open(auth_file),
                                                                            scope: scope)
    goog_retries do
      self.authorization.fetch_access_token!
    end
    self.authorization
  end

  def self.disconnect
    self.authorization = nil
    @drive = @sheets = nil
  end

  def self.authorized?
    self.authorization.present?
  end

  def self.drive
    if @drive.nil?
      raise "No authorizer established" unless self.authorized?
      @drive = Goog::DriveService.new(self.authorization)
    end
    @drive
  end

  def self.sheets
    if @sheets.nil?
      raise "No authorizer established" unless self.authorized?
      @sheets = Goog::SheetsService.new(self.authorization)
    end
    @sheets
  end


end
