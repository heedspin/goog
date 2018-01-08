# https://github.com/google/google-auth-library-ruby
module Goog::Authorizer
  include Goog::Retry
  class << self
    attr_accessor :authorization
  end

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
  end

  def self.established?
    self.authorization.present?
  end
end