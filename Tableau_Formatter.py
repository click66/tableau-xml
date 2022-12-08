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
    
    #Parameter

        #Change original parameter
    for child in tree.findall(".//*[@element='parameter-ctrl-title']/format/[@attr='font-family']"):
        # print(child)
        if child is not None:
            child.set("value", new_font)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter-title']/format/[@attr='color']"))
            log(logging,"filter heading could not be changed")


        #Change manual user updated custom parameter
    for child in tree.findall(".//*[@type-v2='paramctrl']/formatted-text/run/[@fontname]"):
        # print(child)
        if child is not None:
            child.set("fontname", new_font)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter']/format/[@attr='fontcolor']"))
            log(logging,"filter heading could not be changed")

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

    #Parameter

        #Change original parameter
    for child in tree.findall(".//*[@element='parameter-ctrl-title']/format/[@attr='font-family']"):
        # print(child)
        if child is not None:
            child.set("value", new_font)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter-title']/format/[@attr='color']"))
            log(logging,"filter heading could not be changed")


        #Change manual user updated custom parameter
    for child in tree.findall(".//*[@type-v2='paramctrl']/formatted-text/run/[@fontname]"):
        # print(child)
        if child is not None:
            child.set("fontname", new_font)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter']/format/[@attr='fontcolor']"))
            log(logging,"filter heading could not be changed")


    # Get the XML we read in above and write the new element
    return wb


# Filter headers
#############################################################################

def update_filter_headers(wb, new_filter_header_color, new_filter_header_font_size, logging=None):
    # This will replace the filter header format of the workbook

    matched_flag = False
    # Read workbook xml
    tree =wb.xml
    # Update filter heading if found, raise an error if not

#Font colour updates
##################################

    #Filters

        #Change original filters
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


        #Change manual user updated custom filter
    for child in tree.findall(".//*[@element='quick-filter']/format/formatted-text/run/[@fontcolor]"):
        # print(child)
        if child is not None:
            child.set("fontcolor", new_filter_header_color)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter']/format/[@attr='fontcolor']"))
            log(logging,"filter heading could not be changed")
            
    #Parameter

        #Change original parameter
    for child in tree.findall(".//*[@element='parameter-ctrl-title']/format/[@attr='color']"):
        # print(child)
        if child is not None:
            child.set("value", new_filter_header_color)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter-title']/format/[@attr='color']"))
            log(logging,"filter heading could not be changed")


        #Change manual user updated custom parameter
    for child in tree.findall(".//*[@type-v2='paramctrl']/formatted-text/run/[@fontcolor]"):
        # print(child)
        if child is not None:
            child.set("fontcolor", new_filter_header_color)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter']/format/[@attr='fontcolor']"))
            log(logging,"filter heading could not be changed")




#Font bold updates
#Set to always bold will need to add variable if we want to adjust
##################################

    #Filters

        #Change original filters
    for child in tree.findall(".//*[@element='quick-filter-title']/format/[@attr='font-weight']"):
        # print(child)
        if child is not None:
            child.set("value", 'bold')
            # if child.attrib['font-size']:
            #     child.set("value",_new_filter_header_font_size)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter-title']/format/[@attr='font-weight']"))
            log(logging,"filter heading could not be changed")


        #Change manual user updated custom filter
    for child in tree.findall(".//*[@element='quick-filter']/format/formatted-text/run/[@bold]"):
        # print(child)
        if child is not None:
            child.set("bold", 'true')
            # if child.attrib['font-size']:
            #     child.set("value",_new_filter_header_font_size)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter-title']/format/[@attr='font-weight']"))
            log(logging,"filter heading could not be changed")

    #Parameter

        #Change original parameter
    for child in tree.findall(".//*[@element='parameter-ctrl-title']/format/[@attr='font-weight']"):
        # print(child)
        if child is not None:
            child.set("value", 'bold')
            # if child.attrib['font-size']:
            #     child.set("value",_new_filter_header_font_size)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter-title']/format/[@attr='font-weight']"))
            log(logging,"filter heading could not be changed")


        #Change manual user updated custom parameter
    for child in tree.findall(".//*[@type-v2='paramctrl']/formatted-text/run/[@bold]"):
        # print(child)
        if child is not None:
            child.set("bold", 'true')
            # if child.attrib['font-size']:
            #     child.set("value",_new_filter_header_font_size)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter-title']/format/[@attr='font-weight']"))
            log(logging,"filter heading could not be changed")           


#Font Italics updates
#Set to always non italic will need to add variable if we want to adjust
##################################

    #Filters

        #Change original filters
    for child in tree.findall(".//*[@element='quick-filter-title']/format/[@attr='font-style']"):
        # print(child)
        if child is not None:
            child.set("value", '')
            # if child.attrib['font-size']:
            #     child.set("value",_new_filter_header_font_size)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter-title']/format/[@attr='font-weight']"))
            log(logging,"filter heading could not be changed")


        #Change manual user updated custom filter
    for child in tree.findall(".//*[@element='quick-filter']/format/formatted-text/run/[@italic]"):
        # print(child)
        if child is not None:
            child.set("italic", 'false')
            # if child.attrib['font-size']:
            #     child.set("value",_new_filter_header_font_size)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter-title']/format/[@attr='font-weight']"))
            log(logging,"filter heading could not be changed")
    
    #Parameter

        #Change original parameter
    for child in tree.findall(".//*[@element='parameter-ctrl-title']/format/[@attr='font-style']"):
        # print(child)
        if child is not None:
            child.set("value", '')
            # if child.attrib['font-size']:
            #     child.set("value",_new_filter_header_font_size)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter-title']/format/[@attr='font-weight']"))
            log(logging,"filter heading could not be changed")


        #Change manual user updated custom parameter
    for child in tree.findall(".//*[@type-v2='paramctrl']/formatted-text/run/[@italic]"):
        # print(child)
        if child is not None:
            child.set("italic", 'false')
            # if child.attrib['font-size']:
            #     child.set("value",_new_filter_header_font_size)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter-title']/format/[@attr='font-weight']"))
            log(logging,"filter heading could not be changed")   



#Font Underline updates
#Set to always non Underline will need to add variable if we want to adjust
##################################

    #Filters

        #Change original filters
    for child in tree.findall(".//*[@element='quick-filter-title']/format/[@attr='text-decoration']"):
        # print(child)
        if child is not None:
            child.set("value", '')
            # if child.attrib['font-size']:
            #     child.set("value",_new_filter_header_font_size)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter-title']/format/[@attr='font-weight']"))
            log(logging,"filter heading could not be changed")


        #Change manual user updated custom filter
    for child in tree.findall(".//*[@element='quick-filter']/format/formatted-text/run/[@underline]"):
        # print(child)
        if child is not None:
            child.set("underline", 'false')
            # if child.attrib['font-size']:
            #     child.set("value",_new_filter_header_font_size)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter-title']/format/[@attr='font-weight']"))
            log(logging,"filter heading could not be changed")

    #Parameter

       #Change original parameter
    for child in tree.findall(".//*[@element='parameter-ctrl-title']/format/[@attr='text-decoration']"):
        # print(child)
        if child is not None:
            child.set("value", '')
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='parameter-ctrl-title']/format/[@attr='text-decoration']"))
            log(logging,"filter heading could not be changed")


        #Change manual user updated custom parameter
    for child in tree.findall(".//*[@type-v2='paramctrl']/formatted-text/run/[@underline]"):
        # print(child)
        if child is not None:
            child.set("underline", 'false')
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@type-v2='paramctrl']/formatted-text/run/[@underline]"))
            log(logging,"filter heading could not be changed")



#Font Size updates
##################################

    #Filters

        #Change original filters
    for child in tree.findall(".//*[@element='quick-filter-title']/format/[@attr='font-size']"):
        # print(child)
        if child is not None:
            child.set("value", new_filter_header_font_size)
            # if child.attrib['font-size']:
            #     child.set("value",_new_filter_header_font_size)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter-title']/format/[@attr='font-size']"))
            log(logging,"filter heading could not be changed")


            #Change manual user updated custom filter
    for child in tree.findall(".//*[@element='quick-filter']/format/formatted-text/run/[@fontsize]"):
        # print(child)
        if child is not None:
            child.set("fontsize", new_filter_header_font_size)
            # if child.attrib['font-size']:
            #     child.set("value",_new_filter_header_font_size)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='quick-filter-title']/format/[@fontsize']"))
            log(logging,"filter heading could not be changed")

    #Parameter

       #Change original parameter
    for child in tree.findall(".//*[@element='parameter-ctrl-title']/format/[@attr='font-size']"):
        # print(child)
        if child is not None:
            child.set("value", new_filter_header_font_size)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@element='parameter-ctrl-title']/format/[@attr='font-size']"))
            log(logging,"filter heading could not be changed")


        #Change manual user updated custom parameter
    for child in tree.findall(".//*[@type-v2='paramctrl']/formatted-text/run/[@fontsize]"):
        # print(child)
        if child is not None:
            child.set("fontsize", new_filter_header_font_size)
            matched_flag = True
        log(logging,"filter heading has been changed")
    else:
        if matched_flag == False:
            # print(tree.findall(".//*[@type-v2='paramctrl']/formatted-text/run/[@fontsize]"))
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

    for child in tree.findall('zone/[@mode="checklist"]'):
        child.set('mode', 'checkdropdown')
       # log(logging, f"Child {child.get('name')} type changed to checkdropdown")

    return wb


def synchronise_all_filters(wb, logging=None):
    def clean_copy(e):
        f = copy.deepcopy(e)
        f.set('name', f.get('name').partition(' - ')[2])
        return f

    tree = wb.xml

    # Get all zones that contain filters
    filters = tree.findall('.//zone[@type-v2="filter"]')
    filter_zones = list(dict.fromkeys(map(lambda e: e.getparent(), filters)))

    # Get a unique set of filters (by "param" attribute)
    mapped_filters = reduce(lambda acc, e: acc | {e.get('param'): clean_copy(e)}, filters, {})

    # Remove all existing filter elements from workbook
    for filter in filters:
        filter.getparent().remove(filter)
    # for fz in filter_zones:
    #     try:
    #         [fz.remove(e) for e in filters]
    #     except ValueError:
    #         pass

    # Append a copy of all filters to each container zone
    for fz in filter_zones:
        fzs = fz.find('zone-style')
        fz.remove(fzs)
        [fz.append(copy.deepcopy(e)) for e in list(mapped_filters.values())]
        fz.append(fzs)

    # Update the filters to contain the appropriate dashboard names
    dashboards = tree.findall('.//dashboard')
    for d in dashboards:
        for c in d.findall('.//zone[@type-v2="filter"]'):
            c.set('name', d.get('name') + ' - ' + c.get('name'))

    return wb
