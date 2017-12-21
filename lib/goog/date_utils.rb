module Goog::DateUtils
  def date_to_goog(date)
    if date.is_a?(String)
      date = Date.parse(date)
    end
    date - Date.new(1899,12,30)
  end
end