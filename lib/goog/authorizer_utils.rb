# https://github.com/google/google-auth-library-ruby
module Goog::AuthorizerUtils
  # Returns current authorizer
  def authorize_service_account(auth_file:, scope: Google::Apis::DriveV3::AUTH_DRIVE)
    self.current_authorizer = Google::Auth::ServiceAccountCredentials.make_creds(json_key_io: File.open(auth_file),
                                                                                 scope: scope)
    goog_retries do
      self.current_authorizer.fetch_access_token!
    end
    self.current_authorizer
  end

  def self.included base
    base.class_eval do
      attr_accessor :current_authorizer
      include Goog::Retry
    end
  end
end