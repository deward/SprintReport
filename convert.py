# Convert.py: convert json file to a html file with a table in sprint report
# Table example: Template1.orig.htm
#
# Usage:
#   python.exe convert.py jsonfile.json
#
# Output:
#   A htm file wiht the table
#
# Version 0.2               2/1/2016
#

import json
import HTML
from json2html import *


# read json file
with open('data.json') as json_file:
    data = json.load(json_file)

# read header file
with open('Template1-Header.htm') as header_file:
    header_data = header_file.read()


# read foot file
with open('Template1-foot.htm') as foot_file:
    foot_data = foot_file.read()

# ------------------- table format
# global attribute
span_style = '<span style=\'mso-bookmark:_MailOriginal\'></span>'
highlight_background = ';background:#C6D9F1'
# line arritbute
line_row_attribute = 'mso-yfti-irow'
line_height = '15.0pt'
# install
install_width = 59
install_sytle = 'width:44.55pt;border:solid #4F81BD 1.0pt;border-top:none;background:#C6D9F1;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'
install_class = 'MsoNormal'
install_align = 'center'
install_align_style = 'text-align:center'
install_font_style = "font-family:\"Cambria\",serif"
# install_product
install_product_width = 85
install_product_sytle = 'width:64.0pt;border-top:none;border-left:none;border-bottom:solid #4F81BD 1.0pt;border-right:solid #4F81BD 1.0pt;background:#C6D9F1;padding:0in 0in 0in 0in;height:15.0pt'
install_product_class = 'MsoNormal'
install_product_align = 'center'
install_product_align_style = 'text-align:center'
install_product_font_style = "font-family:\"Cambria\",serif"
# item_id
item_id_width = 126
item_id_style = 'width:94.85pt;border-top:none;border-left:none;border-bottom:solid #4F81BD 1.0pt;border-right:solid #4F81BD 1.0pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'
item_id_class = 'MsoNormal'
item_id_font_style = "color:black"
# item_name
item_name_width = 528
item_name_style = 'width:395.65pt;border-top:none;border-left:none;border-bottom:solid #4F81BD 1.0pt;border-right:solid #4F81BD 1.0pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'
item_name_class = 'MsoNormal'
item_name_font_style = "color:black"
# CI_CD
CI_CD_width = 85
CI_CD_sytle = 'border:solid #4F81BD 1.0pt;border-top:none;background:#C6D9F1;padding:0in 0in 0in 0in;height:15.0pt'
CI_CD_class = 'MsoNormal'
CI_CD_align = 'center'
CI_CD_align_style = 'text-align:center'
CI_CD_font_style = "font-family:\"Cambria\",serif;color:black"

background_dark = '#C6D9F1'

# ------------------- generate html file
with open('data.htm', 'w') as data_htm:
    # write header file
    data_htm.write(header_data)

    row_index = 1
    id_style = ''
    name_style = ''

    civil_list_len = len(data['Install']['Civil'])
    plant_list_len = len(data['Install']['Plant'])
    try:
        ia_list_len = len(data['Install']['Install Automation'])
    except (KeyError, IndexError) as e:
        ia_list_len = 0
    try:
        ist_list_len = len(data['Install']['IST'])
    except (KeyError, IndexError) as e:
        ist_list_len = 0
    install_items = civil_list_len + plant_list_len + ia_list_len + ist_list_len

    # write install projects in table - Civil
    row_product = 0
    for item in data['Install']['Civil']:
        if (row_index % 2 == 1):
            id_style = item_id_style + highlight_background
            name_style = item_name_style + highlight_background
        else:
            id_style = item_id_style
            name_style = item_name_style

        if row_product == 0:
            line = '<tr style=\'' + line_row_attribute + ':'+ str(row_index) + ';height:' + str(line_height) +'\'>\n' 
            data_htm.write(line)
            # write 'Install'
            line = '<td width=' + str(install_width) + ' rowspan=' + str(install_items) + ' style=\'' + install_sytle + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + install_class + ' align=' + install_align + ' sytle=\'' + install_align_style + '\'><b><span style=\'' + install_font_style + '\'>Install<o:p></o:p></span></b></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            # write 'Product' - Civil
            line = '<td width=' + str(install_product_width) + ' rowspan=' + str(civil_list_len) + ' style=\'' + install_product_sytle + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + install_product_class + ' align=' + str(install_product_align) + ' sytle=\'' + install_product_align_style + '\'><b><span style=\'' + install_product_font_style + '\'>Civil<o:p></o:p></span></b></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            # write first line
            line = '<td width=' + str(item_id_width) + ' nowrap valign=bottom style=\'' + id_style + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + item_id_class + '><a href=\"' + item[2] + '\">' + item[0] + '</a><span style=\'' + item_id_font_style + '\'><o:p></o:p></span></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            line = '<td width=' + str(item_name_width) + ' nowrap valign=bottom style=\'' + name_style + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + item_name_class + '><span style=\'' + item_name_font_style + '\'>' + item[1] + '<o:p></o:p></span></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            data_htm.write('</tr>\n')
        else:
            # write line
            line = '<tr style=\'' + line_row_attribute + ':'+ str(row_index) + ';height:' + str(line_height) +'\'>\n' 
            data_htm.write(line)
            line = '<td width=' + str(item_id_width) + ' nowrap valign=bottom style=\'' + id_style + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + item_id_class + '><a href=\"' + item[2] + '\">' + item[0] + '</a><span style=\'' + item_id_font_style + '\'><o:p></o:p></span></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            line = '<td width=' + str(item_name_width) + ' nowrap valign=bottom style=\'' + name_style + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + item_name_class + '><span style=\'' + item_name_font_style + '\'>' + item[1] + '<o:p></o:p></span></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            data_htm.write('</tr>\n')
        row_product += 1
        row_index += 1

    # write install projects in table - Plant
    row_product = 0
    for item in data['Install']['Plant']:
        if (row_index % 2 == 1):
            id_style = item_id_style + highlight_background
            name_style = item_name_style + highlight_background
        else:
            id_style = item_id_style
            name_style = item_name_style

        if row_product == 0:
            line = '<tr style=\'' + line_row_attribute + ':'+ str(row_index) + ';height:' + str(line_height) +'\'>\n' 
            data_htm.write(line)
            # write 'Product' - Plant
            line = '<td width=' + str(install_product_width) + ' rowspan=' + str(plant_list_len) + ' style=\'' + install_product_sytle + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + install_product_class + ' align=' + str(install_product_align) + ' sytle=\'' + install_product_align_style + '\'><b><span style=\'' + install_product_font_style + '\'>Plant<o:p></o:p></span></b></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            # write first line
            line = '<td width=' + str(item_id_width) + ' nowrap valign=bottom style=\'' + id_style + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + item_id_class + '><a href=\"' + item[2] + '\">' + item[0] + '</a><span style=\'' + item_id_font_style + '\'><o:p></o:p></span></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            line = '<td width=' + str(item_name_width) + ' nowrap valign=bottom style=\'' + name_style + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + item_name_class + '><span style=\'' + item_name_font_style + '\'>' + item[1] + '<o:p></o:p></span></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            data_htm.write('</tr>\n')
        else:
            # write line
            line = '<tr style=\'' + line_row_attribute + ':'+ str(row_index) + ';height:' + str(line_height) +'\'>\n' 
            data_htm.write(line)
            line = '<td width=' + str(item_id_width) + ' nowrap valign=bottom style=\'' + id_style + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + item_id_class + '><a href=\"' + item[2] + '\">' + item[0] + '</a><span style=\'' + item_id_font_style + '\'><o:p></o:p></span></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            line = '<td width=' + str(item_name_width) + ' nowrap valign=bottom style=\'' + name_style + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + item_name_class + '><span style=\'' + item_name_font_style + '\'>' + item[1] + '<o:p></o:p></span></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            data_htm.write('</tr>\n')
        row_product += 1
        row_index += 1


    # write install projects in table - Install Automation
    row_product = 0
    for item in data['Install']['Install Automation']:
        if (row_index % 2 == 1):
            id_style = item_id_style + highlight_background
            name_style = item_name_style + highlight_background
        else:
            id_style = item_id_style
            name_style = item_name_style

        if row_product == 0:
            line = '<tr style=\'' + line_row_attribute + ':'+ str(row_index) + ';height:' + str(line_height) +'\'>\n' 
            data_htm.write(line)
            # write 'Product' - Install Automation
            line = '<td width=' + str(install_product_width) + ' rowspan=' + str(ia_list_len) + ' style=\'' + install_product_sytle + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + install_product_class + ' align=' + str(install_product_align) + ' sytle=\'' + install_product_align_style + '\'><b><span style=\'' + install_product_font_style + '\'>Install Automation<o:p></o:p></span></b></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            # write first line
            line = '<td width=' + str(item_id_width) + ' nowrap valign=bottom style=\'' + id_style + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + item_id_class + '><a href=\"' + item[2] + '\">' + item[0] + '</a><span style=\'' + item_id_font_style + '\'><o:p></o:p></span></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            line = '<td width=' + str(item_name_width) + ' nowrap valign=bottom style=\'' + name_style + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + item_name_class + '><span style=\'' + item_name_font_style + '\'>' + item[1] + '<o:p></o:p></span></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            data_htm.write('</tr>\n')
        else:
            # write line
            line = '<tr style=\'' + line_row_attribute + ':'+ str(row_index) + ';height:' + str(line_height) +'\'>\n' 
            data_htm.write(line)
            line = '<td width=' + str(item_id_width) + ' nowrap valign=bottom style=\'' + id_style + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + item_id_class + '><a href=\"' + item[2] + '\">' + item[0] + '</a><span style=\'' + item_id_font_style + '\'><o:p></o:p></span></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            line = '<td width=' + str(item_name_width) + ' nowrap valign=bottom style=\'' + name_style + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + item_name_class + '><span style=\'' + item_name_font_style + '\'>' + item[1] + '<o:p></o:p></span></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            data_htm.write('</tr>\n')
        row_product += 1
        row_index += 1

    # write install projects in table - IST
    row_product = 0
    try:
        for item in data['Install']['IST']:
            if (row_index % 2 == 1):
                id_style = item_id_style + highlight_background
                name_style = item_name_style + highlight_background
            else:
                id_style = item_id_style
                name_style = item_name_style

            if row_product == 0:
                line = '<tr style=\'' + line_row_attribute + ':'+ str(row_index) + ';height:' + str(line_height) +'\'>\n' 
                data_htm.write(line)
                # write 'Product' - IST
                line = '<td width=' + str(install_product_width) + ' rowspan=' + str(ist_list_len) + ' style=\'' + install_product_sytle + '\'>\n'
                data_htm.write(line)
                line = '<p class=' + install_product_class + ' align=' + str(install_product_align) + ' sytle=\'' + install_product_align_style + '\'><b><span style=\'' + install_product_font_style + '\'>IST<o:p></o:p></span></b></p>'
                data_htm.write(line)
                data_htm.write('</td>\n')
                # write first line
                line = '<td width=' + str(item_id_width) + ' nowrap valign=bottom style=\'' + id_style + '\'>\n'
                data_htm.write(line)
                line = '<p class=' + item_id_class + '><a href=\"' + item[2] + '\">' + item[0] + '</a><span style=\'' + item_id_font_style + '\'><o:p></o:p></span></p>'
                data_htm.write(line)
                data_htm.write('</td>\n')
                line = '<td width=' + str(item_name_width) + ' nowrap valign=bottom style=\'' + name_style + '\'>\n'
                data_htm.write(line)
                line = '<p class=' + item_name_class + '><span style=\'' + item_name_font_style + '\'>' + item[1] + '<o:p></o:p></span></p>'
                data_htm.write(line)
                data_htm.write('</td>\n')
                data_htm.write('</tr>\n')
            else:
                # write line
                line = '<tr style=\'' + line_row_attribute + ':'+ str(row_index) + ';height:' + str(line_height) +'\'>\n' 
                data_htm.write(line)
                line = '<td width=' + str(item_id_width) + ' nowrap valign=bottom style=\'' + id_style + '\'>\n'
                data_htm.write(line)
                line = '<p class=' + item_id_class + '><a href=\"' + item[2] + '\">' + item[0] + '</a><span style=\'' + item_id_font_style + '\'><o:p></o:p></span></p>'
                data_htm.write(line)
                data_htm.write('</td>\n')
                line = '<td width=' + str(item_name_width) + ' nowrap valign=bottom style=\'' + name_style + '\'>\n'
                data_htm.write(line)
                line = '<p class=' + item_name_class + '><span style=\'' + item_name_font_style + '\'>' + item[1] + '<o:p></o:p></span></p>'
                data_htm.write(line)
                data_htm.write('</td>\n')
                data_htm.write('</tr>\n')
            row_product += 1
            row_index += 1
    except (KeyError) as e:
        print 'No IST items!'

    # write CI/CD
    row_product = 0
    cicd_items = len(data['CI/CD'])
    
    for item in data['CI/CD']:
        if (row_index % 2 == 1):
            id_style = item_id_style + highlight_background
            name_style = item_name_style + highlight_background
        else:
            id_style = item_id_style
            name_style = item_name_style

        if row_product == 0:
            line = '<tr style=\'' + line_row_attribute + ':'+ str(row_index) + ';height:' + str(line_height) +'\'>\n' 
            data_htm.write(line)
            # write 'CI/CD'
            line = '<td colspan=2 rowspan=' + str(cicd_items) + ' style=\'' + CI_CD_sytle + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + CI_CD_class + ' align=' + CI_CD_align + ' sytle=\'' + CI_CD_align_style + '\'><b><span style=\'' + CI_CD_font_style + '\'>CI/CD<o:p></o:p></span></b></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            # write first line
            line = '<td width=' + str(item_id_width) + ' nowrap valign=bottom style=\'' + id_style + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + item_id_class + '><a href=\"' + item[2] + '\">' + item[0] + '</a><span style=\'' + item_id_font_style + '\'><o:p></o:p></span></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            line = '<td width=' + str(item_name_width) + ' nowrap valign=bottom style=\'' + name_style + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + item_name_class + '><span style=\'' + item_name_font_style + '\'>' + item[1] + '<o:p></o:p></span></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            data_htm.write('</tr>\n')
        else:
            # write line
            line = '<tr style=\'' + line_row_attribute + ':'+ str(row_index) + ';height:' + str(line_height) +'\'>\n' 
            data_htm.write(line)
            line = '<td width=' + str(item_id_width) + ' nowrap valign=bottom style=\'' + id_style + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + item_id_class + '><a href=\"' + item[2] + '\">' + item[0] + '</a><span style=\'' + item_id_font_style + '\'><o:p></o:p></span></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            line = '<td width=' + str(item_name_width) + ' nowrap valign=bottom style=\'' + name_style + '\'>\n'
            data_htm.write(line)
            line = '<p class=' + item_name_class + '><span style=\'' + item_name_font_style + '\'>' + item[1] + '<o:p></o:p></span></p>'
            data_htm.write(line)
            data_htm.write('</td>\n')
            data_htm.write('</tr>\n')
        row_product += 1
        row_index += 1

    # write Cover Spy
    row_product = 0
    try:
        coverspy_items = len(data['CoverSpy'])

        for item in data['CoverSpy']:
            if (row_index % 2 == 1):
                id_style = item_id_style + highlight_background
                name_style = item_name_style + highlight_background
            else:
                id_style = item_id_style
                name_style = item_name_style

            if row_product == 0:
                line = '<tr style=\'' + line_row_attribute + ':'+ str(row_index) + ';height:' + str(line_height) +'\'>\n' 
                data_htm.write(line)
                # write 'CoverSpy'
                line = '<td colspan=2 rowspan=' + str(coverspy_items) + ' style=\'' + CI_CD_sytle + '\'>\n'
                data_htm.write(line)
                line = '<p class=' + CI_CD_class + ' align=' + CI_CD_align + ' sytle=\'' + CI_CD_align_style + '\'><b><span style=\'' + CI_CD_font_style + '\'>Cover Spy<o:p></o:p></span></b></p>'
                data_htm.write(line)
                data_htm.write('</td>\n')
                # write first line
                line = '<td width=' + str(item_id_width) + ' nowrap valign=bottom style=\'' + id_style + '\'>\n'
                data_htm.write(line)
                line = '<p class=' + item_id_class + '><a href=\"' + item[2] + '\">' + item[0] + '</a><span style=\'' + item_id_font_style + '\'><o:p></o:p></span></p>'
                data_htm.write(line)
                data_htm.write('</td>\n')
                line = '<td width=' + str(item_name_width) + ' nowrap valign=bottom style=\'' + name_style + '\'>\n'
                data_htm.write(line)
                line = '<p class=' + item_name_class + '><span style=\'' + item_name_font_style + '\'>' + item[1] + '<o:p></o:p></span></p>'
                data_htm.write(line)
                data_htm.write('</td>\n')
                data_htm.write('</tr>\n')
            else:
                # write line
                line = '<tr style=\'' + line_row_attribute + ':'+ str(row_index) + ';height:' + str(line_height) +'\'>\n' 
                data_htm.write(line)
                line = '<td width=' + str(item_id_width) + ' nowrap valign=bottom style=\'' + id_style + '\'>\n'
                data_htm.write(line)
                line = '<p class=' + item_id_class + '><a href=\"' + item[2] + '\">' + item[0] + '</a><span style=\'' + item_id_font_style + '\'><o:p></o:p></span></p>'
                data_htm.write(line)
                data_htm.write('</td>\n')
                line = '<td width=' + str(item_name_width) + ' nowrap valign=bottom style=\'' + name_style + '\'>\n'
                data_htm.write(line)
                line = '<p class=' + item_name_class + '><span style=\'' + item_name_font_style + '\'>' + item[1] + '<o:p></o:p></span></p>'
                data_htm.write(line)
                data_htm.write('</td>\n')
                data_htm.write('</tr>\n')
            row_product += 1
            row_index += 1
    except (KeyError) as e:
        print 'No Cover Spy items!'
    
    data_htm.write(foot_data)

