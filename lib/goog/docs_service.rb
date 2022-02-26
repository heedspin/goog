# https://developers.google.com/drive/v3/web/about-sdk
require 'plutolib/logger_utils'
require 'google/apis/docs_v1'

class Goog::DocsService
  include Plutolib::LoggerUtils
  include Goog::Retry
  
  attr_accessor :drive
  def initialize(authorization)
    @docs = Google::Apis::DocsV1::DocsService.new
    @docs.authorization = authorization
  end

  def get(document_id)
    goog_retries do    
      @docs.get_document(document_id)
    end
  end
end