# https://googleapis.dev/ruby/google-api-client/latest/Google/Apis/DocsV1/DocsService.html#batch_update_document-instance_method
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

  def replace_text(document_id:, mapping:)
    requests = []
    mapping.each do |key, value|
      requests.push({
        replace_all_text: {
          contains_text: {
            text: key,
            match_case: false
          },
          replace_text: value
        }
    })
    end
    # batch_update_document(document_id, batch_update_document_request_object = nil, fields: nil, quota_user: nil, options: nil, &block)
    goog_retries(profile_type: 'Docs#replace_text') do
      @docs.batch_update_document(document_id, 
                                  {requests: requests}, 
                                  {})
    end
    true
  end

end