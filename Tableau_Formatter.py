import copy
from functools import reduce

def log(logging, message):
    if logging:
        logging.info(message)

# Font
#############################################################################

def update_font(wb, new_font, logging=None):
    """
    This will replace the font of the workbook
    """
    matched_flag = False
    # Read workbook xml
    tree = wb.xml
    # Update font if found, raise an error if not
    for child in tree.findall('style/style-rule/[@element="all"]/format'):
        # print(child)
        if child is not None:
            child.set("value", new_font)
            matched_flag = True
            log(logging, 'Font has been changed')

    for child in tree.findall(".//*[@element='label']/format/[@attr='font-family']"):
        # print(child)
        if child is not None:
            child.set("value", new_font)
            matched_flag = True
            log(logging, "Font has been changed")
    else:
        if matched_flag == False:
            log(logging, "Font could not be changed")
    


    return wb
    # # Get the XML we read in above and write the new element
    # tree = ET.ElementTree(root)
     #tree.write(wb.xml, xml_declaration=True, encoding='utf-8')



# Outer Padding
#############################################################################

def update_outerpadding(wb, new_outer_padding, logging=None):
    # This will replace the outer padding of the workbook

    matched_flag = False
    # Read workbook xml
    tree = wb.xml
    # Update outer padding if found, raise an error if not
    # child = tree.findall('dashboards/dashboard/zones/zone/zone-style/format/[@attr="margin"]')

    for child in tree.findall(".//*[@type-v2='filter']/zone-style/format/[@attr='margin']"):
        if child is not None:
            child.set("value", new_outer_padding)
            matched_flag = True
        log(logging, "Outer padding has been changed")
    else:
        if matched_flag == False:
            print(tree.findall(".//*[@type-v2='filter']/zone-style/format/[@attr='margin']"))
            log(logging,"Outer padding could not be changed")

    for child in tree.findall(".//*[@type-v2='filter']/zone-style/format/[@attr='margin-top']"):
        if child is not None:
            child.set("value", new_outer_padding)
            matched_flag = True
        log(logging,"Outer padding has been changed")
    else:
        if matched_flag == False:
            print(tree.findall(".//*[@type-v2='filter']/zone-style/format/[@attr='margin-top']"))
            log(logging,"Outer padding could not be changed")

    # Get the XML we read in above and write the new element
    return wb


# Filter headers
#############################################################################

def update_filter_headers(wb, new_filter_header_color, new_filter_header_font_weight, _new_filter_header_font_size, logging=None):
    # This will replace the filter header format of the workbook

    matched_flag = False
    # Read workbook xml
    tree =wb.xml
    # Update filter heading if found, raise an error if not

    for child in tree.findall(".//*[@element='quick-filter-title']/format/[@attr='color']"):
        # print(child)
        if child is not None:
            child.set("value", new_filter_header_color)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter-title']/format/[@attr='color']"))
            log(logging,"filter heading could not be changed")

    for child in tree.findall(".//*[@element='quick-filter-title']/format/[@attr='font-weight']"):
        # print(child)
        if child is not None:
            child.set("value", new_filter_header_font_weight)
            # if child.attrib['font-size']:
            #     child.set("value",_new_filter_header_font_size)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter-title']/format/[@attr='font-weight']"))
            log(logging,"filter heading could not be changed")

    for child in tree.findall(".//*[@element='quick-filter-title']/format/[@attr='font-size']"):
        # print(child)
        if child is not None:
            child.set("value", _new_filter_header_font_size)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter-title']/format/[@attr='font-size']"))
            log(logging,"filter heading could not be changed")

    # Get the XML we read in above and write the new element
    return wb


def all_filters_show_apply_button(wb, logging=None):
    tree = wb.xml

    for child in tree.findall(".//*[@type-v2='filter']"):
        child.set('show-apply', 'true')

    return wb


def make_multi_filters_checkdropdown(wb, logging=None):
    tree = wb.xml

    for child in tree.findall(".//zone[@type-v2='filter'][@mode='checklist']"):
        child.set('mode', 'checkdropdown')
        log(logging, f"Child {child.get('name')} type changed to checkdropdown")

    return wb


def synchronise_all_filters(wb, logging=None):
    def clean_copy(e):
        f = copy.deepcopy(e)
        f.set('name', f.get('name').partition(' - ')[2])
        return f

    tree = wb.xml

    # Get all zones that contain filters
    filters = tree.findall('.//zone[@type-v2="filter"]')
    filter_zones = list(dict.fromkeys(map(lambda c: c.getparent(), filters)))

    # Get a unique set of filters (by "param" attribute)
    mapped_filters = reduce(lambda acc, e: acc | {e.get('param'): e}, filters, {})

    # Remove all existing filter elements from workbook
    for fz in filter_zones:
        try:
            [fz.remove(e) for e in filters]
        except ValueError:
            pass

    # Append a copy of all filters to each container zone
    for fz in filter_zones:
        [fz.append(clean_copy(e)) for e in list(mapped_filters.values())]

    # Update the filters to contain the appropriate dashboard names
    dashboards = tree.findall('.//dashboard')
    for d in dashboards:
        for c in d.findall('.//zone[@type-v2="filter"]'):
            c.set('name', d.get('name') + ' - ' + c.get('name'))

    return wb
