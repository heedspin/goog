# https://googleapis.dev/ruby/google-api-client/latest/Google/Apis/DocsV1/DocsService.html#batch_update_document-instance_method
require 'plutolib/logger_utils'
require 'google/apis/docs_v1'

class Goog::DocsService
  include Plutolib::LoggerUtils
  include Goog::Retry
  
  attr_accessor :docs
  def initialize(authorization)
    @docs = Google::Apis::DocsV1::DocsService.new
    @docs.authorization = authorization
  end

  def get(document_id)
    goog_retries do    
      @docs.get_document(document_id)
    end
  end

  def replace_doc(source_file_id:, destination_file_id:)
    source_io = StringIO.new
    mime_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    Goog::Services.drive.drive.export_file(source_file_id, 
                                           mime_type,
                                           download_dest: source_io)
    goog_retries do
      file_id = Goog::Services.drive.drive.update_file(destination_file_id,
                                                       upload_source: source_io,
                                                       content_type: mime_type)
    end
    true
  end  

  def replace_text(document_id:, mapping:)
    requests = []
    # If you go be length desc solves the LOAN_AMOUNT, LOAN_AMOUNT_ENGLISH => $100,000_ENGLISH problem.
    mapping.keys.sort_by(&:length).reverse.each do |key|
      value = mapping[key]
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