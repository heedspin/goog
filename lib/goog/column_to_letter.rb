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
end