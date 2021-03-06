module Goog
  module Retry
    GOOG_MAX_RETRIES=10

    def goog_retries(args={}, &block)
      self.class.goog_retries(args, &block)
    end

    def self.included base
      base.class_eval do
        def self.goog_retry_or_raise(attempts, exception)
          if attempts == Goog::Retry::GOOG_MAX_RETRIES
            log_error "Giving up..."
            raise exception
          else
            delay = attempts * 5
            log_error "Attempt #{attempts}.  Retrying after #{delay} seconds..."
            sleep(delay)
          end
        end

        def self.goog_retries(args={}, &block)
          if args[:profile_type] and Goog::Services.session.try(:profiling_enabled?)
            Goog::Services.session.profile_event(args[:profile_type], args[:profile_name])
          end
          attempts = 0
          while ((attempts += 1) <= Goog::Retry::GOOG_MAX_RETRIES)
            begin
              return yield
            # rescue HTTPClient::ReceiveTimeoutError
            rescue Google::Apis::TransmissionError => exception
              log_error "Transmission Error: #{exception.message}."
              goog_retry_or_raise(attempts, exception)
            rescue Google::Apis::ServerError => exception
              log_error "Server Error: #{exception.message}."
              goog_retry_or_raise(attempts, exception)          
            rescue Google::Apis::RateLimitError => exception
              # pp JSON.parse(err.body)
              # {"error"=>
              #   {"code"=>429,
              #    "message"=>
              #     "Insufficient tokens for quota 'WriteGroup' and limit 'USER-100s' of service 'sheets.googleapis.com' for consumer 'project_number:122247535077'.",
              #    "errors"=>
              #     [{"message"=>
              #        "Insufficient tokens for quota 'WriteGroup' and limit 'USER-100s' of service 'sheets.googleapis.com' for consumer 'project_number:122247535077'.",
              #       "domain"=>"global",
              #       "reason"=>"rateLimitExceeded"}],
              #    "status"=>"RESOURCE_EXHAUSTED"}}
              # {"error"=>{"code"=>429, "message"=>"Insufficient tokens for quota 'WriteGroup' and limit 'USER-100s' of service 'sheets.googleapis.com' for consumer 'project_number:122247535077'.", "errors"=>[{"message"=>"Insufficient tokens for quota 'WriteGroup' and limit 'USER-100s' of service 'sheets.googleapis.com' for consumer 'project_number:122247535077'.", "domain"=>"global", "reason"=>"rateLimitExceeded"}], "status"=>"RESOURCE_EXHAUSTED"}}
              error = JSON.parse(exception.body)['error']
              log_error "Rate Limit Error: #{error['message']}."
              goog_retry_or_raise(attempts, exception)
            rescue StandardError => error
              log_error "Caught unexpected error:", error
              raise $!
            end
          end
        end
      end
    end
  end
end