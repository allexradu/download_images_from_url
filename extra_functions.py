str = ''
ind = 0
number = 0


def value_key(cell_letter, cell_number):
    """ Returning a Cell Row combination to use as a key"""
    return '{cell_letter}{cell_number}'.format(cell_letter = cell_letter,
                                               cell_number = cell_number)
