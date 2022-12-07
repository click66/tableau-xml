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
    # tree.write(wbname, xml_declaration=True, encoding='utf-8')


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
        if child is not NONE:
            child.set("value", new_outer_padding)
            matched_flag = True
        log(logging, "Outer padding has been changed")
    else:
        if matched_flag == False:
            print(tree.findall(".//*[@type-v2='filter']/zone-style/format/[@attr='margin']"))
            log(logging,"Outer padding could not be changed")

    for child in tree.findall(".//*[@type-v2='filter']/zone-style/format/[@attr='margin-top']"):
        if child is not NONE:
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
