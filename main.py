'''
This app allows the user to select a MS Outlook message from McMaster and it
    copies the relevant data to the clipboard for pasting into Trello in Markup
    Quite a few bits of clunk can be taken out of this: I'm aware but while it's working
    I'm going to leave it alone :P
Daniel Woodson 20200527
Version 1.1
Version Note: This version contains the part description and unit cost
'''

import pyperclip
import utilities

if __name__ == '__main__':
    message = utilities.pick_message()
    contents = str(message.body)
    links = utilities.find_links(contents)
    item_numbers = utilities.count_items(links)
    part_numbers = utilities.get_part_numbers(links)
    qtys = utilities.get_quantities(contents, part_numbers)
    descriptions = utilities.get_description(contents)
    costs = utilities.get_cost(contents, part_numbers)
    details = utilities.get_details(contents, part_numbers)
    pyperclip.copy(utilities.build_message(links, item_numbers, part_numbers, qtys, descriptions, costs, details))

