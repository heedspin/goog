module Goog::ColumnToLetter
  # https://stackoverflow.com/questions/21229180/convert-column-index-into-corresponding-column-letter
  def column_to_letter(column)
    letter = ''
    while (column > 0)
      temp = (column - 1) % 26
      letter = (temp + 65).chr + letter
      column = (column - temp - 1) / 26
    end
    letter
  end
  def letter_to_column(letter)
    column = 0
    length = letter.length
    for i in 0..(length-1)
      column += (letter[i].ord - 64) * (26 ** (length - i - 1))
    end
    column
  end

  # Returns [Sheet, Range1, Range2]
  def get_range_parts(range)
    if range =~ /([' a-zA-Z]+!)?([A-Z]+\d+):([A-Z]+\d+)/
      sheet = $1
      range1 = $2
      range2 = $3
      if sheet
        sheet = sheet[0..sheet.size-2]
        if sheet =~ /'(.+)'/
          sheet = $1
        end
      end
      return [sheet, range1, range2]
    end
    nil
  end

end