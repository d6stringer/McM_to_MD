import re
from outlook_msg import Message
from tkinter import Tk
from tkinter import filedialog


def pick_message():
    try:

        root = Tk()
        root.withdraw()
        root.call('wm', 'attributes', '.', '-topmost', True)
        filename = filedialog.askopenfilename(filetypes = [('Outlook Messages', '*.msg')])
        print(filename)
        with open(filename) as msg_file:
            msg = Message(msg_file)
        return msg
    except:
        pass


def build_message(links, item_numbers, part_numbers, qtys, descriptions, costs, details):
    prnt_string = ''
    for i in descriptions:
        max_len_desc = 0
        if len(i) > max_len_desc:
            max_len_desc = len(i)

    for i in range(len(item_numbers)):
        part_numbers[i] = part_numbers[i].rjust(8, ' ')
        descriptions[i] = descriptions[i].ljust(max_len_desc, ' ')
        prnt_string = prnt_string + str(item_numbers[i]) + ') [' + part_numbers[i] +']' + '(' + links[i] + ') ' + 'QTY: ' + qtys[i] + descriptions[i] + costs[i] + 'Each\n' + details[i] + '\n\n'
    return prnt_string

def find_links(contents):
    # find all the links and remove the last 2 (bc they are garbage
    links = re.findall('<(.*)>', contents)
    links = links[:-2]
    #print(links)
    return links

def count_items(list):
    # create list with number of items (silly I know but will make my life easier later)
    item_numbers = []
    item_numbers.extend(range(1, len(list) + 1))
    #print(item_numbers)
    return item_numbers

def get_part_numbers(links):
    # create a list of each part number
    part_numbers = []
    for i in range(len(links)):
        part = str(links[i])
        part = part[part.index('g/') + 2:]
        part_numbers.append(part)
    #print(part_numbers)
    return part_numbers

def get_quantities(contents, part_numbers):
    # get quantities from email
    split_contents = contents.splitlines()
    qty_indicies = []
    for line in split_contents:
        for i in part_numbers:
            if i in line:
                qty_indicies.append(split_contents.index(line) + 1)

    qty_indicies = qty_indicies[1::2]
    qtys = []
    for i in range(len(qty_indicies)):
        qtys.append(split_contents[qty_indicies[i]])
    #print(qtys)
    return qtys

def get_cost(contents, part_numbers):
    # get quantities from email
    split_contents = contents.splitlines()
    cost_indicies = []
    for line in split_contents:
        for i in part_numbers:
            if i in line:
                cost_indicies.append(split_contents.index(line) + 3)

    cost_indicies = cost_indicies[1::2]
    costs = []
    for i in range(len(cost_indicies)):
        costs.append(split_contents[cost_indicies[i]])
    #print(qtys)
    return costs

def get_description(contents):
    split_contents = contents.splitlines()
    description_lines = []
    links = find_links(contents)
    for line in split_contents:
        for i in links:
            if i in line:
                description_lines.append(line)

    descriptions = []
    for i in description_lines:
        desc = str(re.findall('\\t(.*)<', i ))
        desc = desc[2:-2]
        descriptions.append(desc)

    return descriptions

def get_details(contents, part_numbers):
    # get details of the part from email
    split_contents = contents.splitlines()
    detail_indicies = []
    for line in split_contents:
        for i in part_numbers:
            if i in line:
                detail_indicies.append(split_contents.index(line) - 1)

    detail_indicies = detail_indicies[1::2]
    details = []
    for i in range(len(detail_indicies)):
        details.append(split_contents[detail_indicies[i]])
    print(details)
    return details